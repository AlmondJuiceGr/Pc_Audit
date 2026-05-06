import argparse
import getpass
import json
import math
import os
import platform
import shutil
import socket
import struct
import subprocess
import sys
import tempfile
import time
import wave
from urllib import error as urlerror
from urllib import request as urlrequest


DEFAULT_SERVER = os.environ.get("JOB_SERVER_URL", "http://127.0.0.1:8765")
DEFAULT_NAME = os.environ.get("CLIENT_NAME", socket.gethostname())
DEFAULT_ROLE = os.environ.get("CLIENT_ROLE", "Client")
DEFAULT_INTERVAL = int(os.environ.get("CLIENT_HEARTBEAT_SECONDS", "60"))


def platform_label():
    return f"{platform.system()} {platform.release()} ({platform.machine()})"


def run_command(command, timeout=6):
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        kwargs = {
            "capture_output": True,
            "text": True,
            "errors": "replace",
            "timeout": timeout,
        }
        if os.name == "nt":
            kwargs["creationflags"] = flags
        result = subprocess.run(command, **kwargs)
    except Exception:
        return ""
    return result.stdout.strip() if result.returncode == 0 else ""


def run_command_result(command, timeout=6):
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        kwargs = {
            "capture_output": True,
            "text": True,
            "errors": "replace",
            "timeout": timeout,
        }
        if os.name == "nt":
            kwargs["creationflags"] = flags
        result = subprocess.run(command, **kwargs)
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as exc:
        return None, "", str(exc)


def powershell_value(script):
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "$ErrorActionPreference='SilentlyContinue'; " + script,
    ]
    value = run_command(command)
    if value:
        return value
    command[0] = "pwsh"
    return run_command(command)


def first_line(value):
    for line in str(value).splitlines():
        line = line.strip()
        if line:
            return line
    return ""


def read_text_file(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as handle:
            return handle.read().strip()
    except OSError:
        return ""


def read_binary_file(path):
    try:
        with open(path, "rb") as handle:
            return handle.read()
    except OSError:
        return b""


def format_gb(bytes_value):
    try:
        gb = int(float(bytes_value)) / (1024 ** 3)
    except (TypeError, ValueError):
        return ""
    rounded = round(gb)
    return f"{rounded} GB" if rounded > 0 else ""


def chassis_label(value):
    try:
        chassis = int(str(value).strip().split()[0])
    except (TypeError, ValueError, IndexError):
        return ""
    if chassis in {8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32}:
        return "LAPTOP"
    if chassis in {3, 4, 5, 6, 7, 15, 16}:
        return "DESKTOP"
    if chassis in {13, 17, 23, 24}:
        return "TABLET"
    return ""


def infer_screen_type_from_text(text):
    haystack = f" {text} ".lower()
    if "amoled" in haystack:
        return "AMOLED"
    if "oled" in haystack:
        return "OLED"
    if "mini led" in haystack or "miniled" in haystack:
        return "Mini LED"
    if "ips" in haystack:
        return "IPS LCD"
    if " va " in haystack or "mva" in haystack or "pva" in haystack:
        return "VA LCD"
    if " tn " in haystack or "twisted nematic" in haystack:
        return "TN LCD"
    if "lcd" in haystack or "led" in haystack:
        return "LCD"
    return "Unknown"


def edid_descriptor_text(edid_bytes):
    if len(edid_bytes) < 128:
        return ""
    chunks = []
    for offset in range(54, 126, 18):
        block = edid_bytes[offset:offset + 18]
        if len(block) == 18 and block[:3] == b"\x00\x00\x00" and block[3] in {0xfc, 0xfe, 0xff}:
            text = block[5:18].decode("ascii", errors="ignore").strip().replace("\n", " ")
            if text:
                chunks.append(text)
    return " ".join(chunks)


def windows_monitor_text():
    monitor_name = powershell_value(
        "Get-CimInstance -Namespace root/wmi -ClassName WmiMonitorID | "
        "ForEach-Object { "
        "$name=($_.UserFriendlyName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''; "
        "$serial=($_.SerialNumberID | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''; "
        "$name + ' ' + $serial }"
    )
    desktop_monitor = powershell_value(
        "Get-CimInstance Win32_DesktopMonitor | ForEach-Object { $_.Name + ' ' + $_.MonitorType }"
    )
    return " ".join([monitor_name, desktop_monitor])


def linux_monitor_text():
    parts = []
    drm_path = "/sys/class/drm"
    if os.path.isdir(drm_path):
        for name in os.listdir(drm_path):
            edid = read_binary_file(os.path.join(drm_path, name, "edid"))
            text = edid_descriptor_text(edid)
            if text:
                parts.append(text)
    xrandr = run_command(["xrandr", "--verbose"], timeout=4)
    if xrandr:
        parts.append(xrandr)
    return " ".join(parts)


def linux_disk_info():
    output = run_command(["lsblk", "-bdno", "NAME,SIZE,ROTA,TYPE"])
    if output:
        for line in output.splitlines():
            parts = line.split()
            if len(parts) >= 4 and parts[3] == "disk":
                return ("SATA HDD" if parts[2] == "1" else "SATA SSD"), format_gb(parts[1])
    if os.path.isdir("/sys/block"):
        for name in os.listdir("/sys/block"):
            if name.startswith(("loop", "ram", "fd")):
                continue
            size = read_text_file(f"/sys/block/{name}/size")
            rotational = read_text_file(f"/sys/block/{name}/queue/rotational")
            if size:
                try:
                    return ("SATA HDD" if rotational == "1" else "SATA SSD"), format_gb(int(size) * 512)
                except ValueError:
                    continue
    return "", ""


def status_item(name, status, details=""):
    return {"name": name, "status": status, "details": details}


def summarize_status(items):
    summary = {"OK": 0, "WARN": 0, "FAIL": 0, "UNKNOWN": 0}
    for item in items:
        status = item.get("status", "UNKNOWN")
        summary[status] = summary.get(status, 0) + 1
    return summary


def windows_device_diagnostics():
    items = []

    optical = first_line(
        powershell_value(
            "Get-CimInstance Win32_CDROMDrive | Select-Object -First 1 | "
            "ForEach-Object { $_.Name + '|' + $_.MediaLoaded }"
        )
    )
    if optical:
        name, _, media_loaded = optical.partition("|")
        media_text = "media loaded" if media_loaded.strip().lower() == "true" else "no media detected"
        items.append(status_item("Optical drive", "OK", f"{name.strip()} ({media_text})"))
    else:
        items.append(status_item("Optical drive", "UNKNOWN", "No optical drive reported by Windows"))

    usb_count = first_line(
        powershell_value("(Get-CimInstance Win32_USBControllerDevice | Measure-Object).Count")
    )
    usb_details = first_line(
        powershell_value(
            "Get-CimInstance Win32_USBController | Select-Object -First 3 -ExpandProperty Name"
        )
    )
    if usb_count and usb_count != "0":
        items.append(status_item("USB", "OK", f"{usb_count} USB device/controller associations; {usb_details}"))
    else:
        items.append(status_item("USB", "WARN", "No USB devices/controllers reported"))

    lan = first_line(
        powershell_value(
            "Get-CimInstance Win32_NetworkAdapter | "
            "Where-Object { $_.PhysicalAdapter -and $_.NetConnectionID -match 'Ethernet|LAN' } | "
            "Select-Object -First 1 | ForEach-Object { $_.Name + '|' + $_.NetConnectionStatus + '|' + $_.Speed }"
        )
    )
    if lan:
        name, _, rest = lan.partition("|")
        status_code, _, speed = rest.partition("|")
        status = "OK" if status_code.strip() == "2" else "WARN"
        items.append(status_item("LAN", status, f"{name.strip()} status={status_code.strip()} speed={speed.strip()}"))
    else:
        items.append(status_item("LAN", "UNKNOWN", "No physical LAN adapter found"))

    wifi = first_line(
        powershell_value(
            "Get-CimInstance Win32_NetworkAdapter | "
            "Where-Object { $_.PhysicalAdapter -and $_.NetConnectionID -match 'Wi-Fi|Wireless|WLAN' } | "
            "Select-Object -First 1 | ForEach-Object { $_.Name + '|' + $_.NetConnectionStatus }"
        )
    )
    if wifi:
        items.append(status_item("Wi-Fi", "OK" if "|2" in wifi else "WARN", wifi))
    else:
        items.append(status_item("Wi-Fi", "UNKNOWN", "No Wi-Fi adapter found"))

    disk = first_line(
        powershell_value(
            "Get-CimInstance Win32_DiskDrive | Select-Object -First 1 | "
            "ForEach-Object { $_.Model + '|' + $_.Status }"
        )
    )
    if disk:
        items.append(status_item("Disk", "OK" if disk.lower().endswith("|ok") else "WARN", disk))
    else:
        items.append(status_item("Disk", "UNKNOWN", "No disk status reported"))

    battery = first_line(
        powershell_value(
            "Get-CimInstance Win32_Battery | Select-Object -First 1 | "
            "ForEach-Object { $_.Name + '|' + $_.EstimatedChargeRemaining + '%' }"
        )
    )
    if battery:
        items.append(status_item("Battery", "OK", battery))
    else:
        items.append(status_item("Battery", "UNKNOWN", "No battery reported"))

    camera = first_line(
        powershell_value(
            "Get-CimInstance Win32_PnPEntity | "
            "Where-Object { $_.Name -match 'Camera|Webcam' } | Select-Object -First 1 -ExpandProperty Name"
        )
    )
    items.append(status_item("Camera", "OK" if camera else "UNKNOWN", camera or "No camera reported"))

    audio = first_line(
        powershell_value(
            "Get-CimInstance Win32_SoundDevice | Select-Object -First 1 | "
            "ForEach-Object { $_.Name + '|' + $_.Status }"
        )
    )
    items.append(status_item("Audio", "OK" if audio.lower().endswith("|ok") else "UNKNOWN", audio or "No audio device reported"))

    return {"items": items, "summary": summarize_status(items)}


def linux_device_diagnostics():
    items = []

    optical = ""
    cdrom_info = read_text_file("/proc/sys/dev/cdrom/info")
    for line in cdrom_info.splitlines():
        if line.startswith("drive name:"):
            optical = line.split(":", 1)[1].strip()
            break
    if optical:
        items.append(status_item("Optical drive", "OK", optical))
    else:
        items.append(status_item("Optical drive", "UNKNOWN", "No optical drive reported by Linux"))

    lsusb = run_command(["lsusb"])
    if lsusb:
        usb_lines = [line for line in lsusb.splitlines() if line.strip()]
        items.append(status_item("USB", "OK", f"{len(usb_lines)} USB device(s) visible"))
    elif os.path.isdir("/sys/bus/usb/devices"):
        usb_devices = [name for name in os.listdir("/sys/bus/usb/devices") if ":" not in name]
        status = "OK" if usb_devices else "WARN"
        items.append(status_item("USB", status, f"{len(usb_devices)} USB sysfs device(s) visible"))
    else:
        items.append(status_item("USB", "UNKNOWN", "No USB data available"))

    lan_details = []
    net_path = "/sys/class/net"
    if os.path.isdir(net_path):
        for iface in os.listdir(net_path):
            if iface == "lo":
                continue
            operstate = read_text_file(os.path.join(net_path, iface, "operstate"))
            address = read_text_file(os.path.join(net_path, iface, "address"))
            speed = read_text_file(os.path.join(net_path, iface, "speed"))
            lan_details.append(f"{iface}: {operstate or 'unknown'} {speed + 'Mbps' if speed else ''} {address}")
    if lan_details:
        status = "OK" if any("up" in item for item in lan_details) else "WARN"
        items.append(status_item("Network adapters", status, "; ".join(lan_details[:4])))
    else:
        items.append(status_item("Network adapters", "UNKNOWN", "No network interfaces found"))

    disk = run_command(["lsblk", "-ndo", "NAME,TYPE,SIZE,MODEL"])
    disk_lines = [line.strip() for line in disk.splitlines() if " disk " in f" {line} "]
    if disk_lines:
        items.append(status_item("Disk", "OK", disk_lines[0]))
    else:
        items.append(status_item("Disk", "UNKNOWN", "No disk details reported"))

    battery_path = "/sys/class/power_supply"
    batteries = []
    if os.path.isdir(battery_path):
        for name in os.listdir(battery_path):
            if name.startswith("BAT"):
                capacity = read_text_file(os.path.join(battery_path, name, "capacity"))
                status = read_text_file(os.path.join(battery_path, name, "status"))
                batteries.append(f"{name}: {capacity}% {status}".strip())
    if batteries:
        items.append(status_item("Battery", "OK", "; ".join(batteries)))
    else:
        items.append(status_item("Battery", "UNKNOWN", "No battery reported"))

    video_devices = []
    if os.path.isdir("/dev"):
        video_devices = [name for name in os.listdir("/dev") if name.startswith("video")]
    items.append(status_item("Camera", "OK" if video_devices else "UNKNOWN", ", ".join(video_devices) or "No /dev/video devices"))

    audio_cards = read_text_file("/proc/asound/cards")
    items.append(status_item("Audio", "OK" if audio_cards else "UNKNOWN", first_line(audio_cards) or "No ALSA cards reported"))

    return {"items": items, "summary": summarize_status(items)}


def collect_device_diagnostics():
    system = platform.system().lower()
    try:
        if system.startswith("win"):
            return windows_device_diagnostics()
        if system == "linux":
            return linux_device_diagnostics()
    except Exception as exc:
        return {"items": [status_item("Diagnostics", "FAIL", str(exc))], "summary": {"FAIL": 1}}
    return {"items": [status_item("Diagnostics", "UNKNOWN", f"Unsupported OS: {platform.system()}")], "summary": {"UNKNOWN": 1}}


def classify_health_text(text):
    lower = str(text).lower()
    if not lower.strip():
        return "UNKNOWN"
    fail_words = ("failed", "failure", "critical", "unhealthy", "bad", "pred fail")
    warn_words = ("warning", "warn", "degraded", "caution", "error", "unknown", "lost communication")
    ok_words = ("healthy", "ok", "passed", "normal", "no errors")
    if any(word in lower for word in fail_words):
        return "FAIL"
    if any(word in lower for word in warn_words):
        return "WARN"
    if any(word in lower for word in ok_words):
        return "OK"
    return "UNKNOWN"


def pc_auditor_disk_health_test():
    system = platform.system().lower()
    lines = []

    if system.startswith("win"):
        details = powershell_value(
            "Get-PhysicalDisk | Select-Object -First 8 | "
            "ForEach-Object { $_.FriendlyName + '|' + $_.HealthStatus + '|' + "
            "($_.OperationalStatus -join ',') + '|' + $_.MediaType }"
        )
        lines = [line.strip() for line in details.splitlines() if line.strip()]
        if not lines:
            fallback = powershell_value(
                "Get-CimInstance Win32_DiskDrive | Select-Object -First 8 | "
                "ForEach-Object { $_.Model + '|' + $_.Status }"
            )
            lines = [line.strip() for line in fallback.splitlines() if line.strip()]

    elif system == "linux":
        smartctl = shutil.which("smartctl")
        if smartctl:
            disk_names = []
            lsblk = run_command(["lsblk", "-ndo", "NAME,TYPE"], timeout=6)
            for line in lsblk.splitlines():
                parts = line.split()
                if len(parts) >= 2 and parts[1] == "disk":
                    disk_names.append(parts[0])
            for name in disk_names[:8]:
                rc, stdout, stderr = run_command_result([smartctl, "-H", f"/dev/{name}"], timeout=12)
                combined = "\n".join(part for part in (stdout, stderr) if part).strip()
                if combined:
                    health_line = first_line(
                        "\n".join(line for line in combined.splitlines() if "health" in line.lower() or "passed" in line.lower())
                    )
                    lines.append(f"/dev/{name}|{health_line or first_line(combined)}|rc={rc}")
        if not lines:
            disk = run_command(["lsblk", "-ndo", "NAME,TYPE,SIZE,MODEL"], timeout=6)
            lines = [line.strip() for line in disk.splitlines() if " disk " in f" {line} "]

    if not lines:
        return status_item("Test disk health", "UNKNOWN", "No disk health information available")

    status = "OK"
    for line in lines:
        line_status = classify_health_text(line)
        if line_status == "FAIL":
            status = "FAIL"
            break
        if line_status in {"WARN", "UNKNOWN"} and status == "OK":
            status = "WARN"

    return status_item("Test disk health", status, "; ".join(lines[:8]))


def pc_auditor_server_packet_test(server_url, attempts=3, timeout=8):
    attempts = max(1, min(int(attempts or 3), 20))
    successes = 0
    failures = []
    latencies = []
    url = server_url.rstrip("/") + "/health"

    for index in range(1, attempts + 1):
        start = time.perf_counter()
        try:
            req = urlrequest.Request(url, method="GET", headers={"Accept": "application/json"})
            with urlrequest.urlopen(req, timeout=timeout) as response:
                response.read(1024)
                code = response.getcode()
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            if 200 <= code < 300:
                successes += 1
                latencies.append(elapsed_ms)
            else:
                failures.append(f"attempt {index}: HTTP {code}")
        except Exception as exc:
            failures.append(f"attempt {index}: {exc}")

    if successes == attempts:
        status = "OK"
    elif successes:
        status = "WARN"
    else:
        status = "FAIL"

    if latencies:
        details = f"{successes}/{attempts} packets OK; latency ms min/avg/max={min(latencies)}/{sum(latencies)//len(latencies)}/{max(latencies)}"
    else:
        details = f"0/{attempts} packets OK"
    if failures:
        details += " | " + "; ".join(failures[:4])
    return status_item("Test server packets", status, details)


def normalize_test_path(path):
    path = str(path or "").strip().strip('"').strip("'")
    if os.name == "nt" and len(path) == 2 and path[1] == ":":
        return path + "\\"
    return os.path.expanduser(path)


def windows_logical_drive_candidates(drive_type):
    output = powershell_value(
        f"Get-CimInstance Win32_LogicalDisk | Where-Object {{ $_.DriveType -eq {drive_type} }} | "
        "ForEach-Object { $_.DeviceID }"
    )
    return [normalize_test_path(line) for line in output.splitlines() if line.strip()]


def linux_media_candidates():
    candidates = []
    user_names = [os.environ.get("USER", ""), os.environ.get("USERNAME", "")]
    bases = ["/media", "/mnt", "/cdrom"]
    for user_name in user_names:
        if user_name:
            bases.extend([f"/media/{user_name}", f"/run/media/{user_name}"])

    for base in bases:
        if not os.path.isdir(base):
            continue
        candidates.append(base)
        try:
            for name in os.listdir(base):
                path = os.path.join(base, name)
                if os.path.isdir(path):
                    candidates.append(path)
        except OSError:
            continue

    seen = set()
    unique = []
    for path in candidates:
        if path not in seen:
            seen.add(path)
            unique.append(path)
    return unique


def test_path_candidates(kind):
    system = platform.system().lower()
    if system.startswith("win"):
        if kind == "optical":
            candidates = windows_logical_drive_candidates(5)
            cdrom = powershell_value(
                "Get-CimInstance Win32_CDROMDrive | Where-Object { $_.Drive } | ForEach-Object { $_.Drive }"
            )
            candidates.extend(normalize_test_path(line) for line in cdrom.splitlines() if line.strip())
            return list(dict.fromkeys(candidates))
        if kind == "removable":
            return windows_logical_drive_candidates(2)
    if system == "linux":
        return linux_media_candidates()
    return []


def choose_test_path(label, provided_path, candidates, prompt, non_interactive=False):
    if provided_path:
        return normalize_test_path(provided_path)

    existing = [path for path in candidates if os.path.exists(normalize_test_path(path))]
    if non_interactive:
        return normalize_test_path(existing[0]) if existing else ""

    print(f"\n{label}")
    print(prompt)
    if existing:
        print("Detected path(s):")
        for path in existing[:8]:
            print(f"  {path}")
    answer = input("Path/drive [Enter=first detected, s=skip]: ").strip()
    if answer.lower() in {"s", "skip"}:
        return ""
    if not answer and existing:
        return normalize_test_path(existing[0])
    return normalize_test_path(answer)


def read_sample_from_path(path, test_name, max_files=32, max_bytes=16 * 1024 * 1024):
    path = normalize_test_path(path)
    if not path:
        return status_item(test_name, "UNKNOWN", "Skipped; no media path was provided")
    if not os.path.exists(path):
        return status_item(test_name, "FAIL", f"Path not found: {path}")

    files = []
    if os.path.isfile(path):
        files = [path]
    else:
        try:
            for root, _, names in os.walk(path):
                for name in names:
                    files.append(os.path.join(root, name))
                    if len(files) >= max_files:
                        break
                if len(files) >= max_files:
                    break
        except OSError as exc:
            return status_item(test_name, "FAIL", f"Could not scan {path}: {exc}")

    if not files:
        return status_item(test_name, "WARN", f"No readable files found in {path}")

    bytes_read = 0
    files_read = 0
    errors = []
    for file_path in files:
        try:
            with open(file_path, "rb") as handle:
                remaining = max(1, max_bytes - bytes_read)
                chunk = handle.read(min(1024 * 1024, remaining))
                bytes_read += len(chunk)
            files_read += 1
        except OSError as exc:
            errors.append(f"{os.path.basename(file_path)}: {exc}")
        if bytes_read >= max_bytes:
            break

    if files_read and not errors:
        status = "OK"
    elif files_read:
        status = "WARN"
    else:
        status = "FAIL"

    sample_names = ", ".join(os.path.basename(name) for name in files[:5])
    details = f"Read {files_read}/{len(files)} sampled file(s), {bytes_read} byte(s) from {path}"
    if sample_names:
        details += f" | sample: {sample_names}"
    if errors:
        details += " | errors: " + "; ".join(errors[:3])
    return status_item(test_name, status, details)


def pc_auditor_optical_read_test(args):
    path = choose_test_path(
        "Optical media read test",
        args.optical_test_path,
        test_path_candidates("optical"),
        "Insert a known-good finalized CD/DVD, wait until it is mounted, then provide the drive/path.",
        non_interactive=args.non_interactive_tests,
    )
    return read_sample_from_path(path, "Test optical read")


def pc_auditor_removable_read_test(args):
    path = choose_test_path(
        "USB/removable storage read test",
        args.usb_test_path,
        test_path_candidates("removable"),
        "Insert a known-good USB flash drive or SD card, wait until it appears, then provide the drive/path.",
        non_interactive=args.non_interactive_tests,
    )
    return read_sample_from_path(path, "Test removable storage read")


def write_stereo_channel_wav(path):
    sample_rate = 44100
    frames = bytearray()

    def add_tone(duration, left_gain, right_gain, frequency=740):
        sample_count = int(sample_rate * duration)
        for index in range(sample_count):
            value = math.sin(2 * math.pi * frequency * index / sample_rate)
            left = int(32767 * 0.55 * left_gain * value)
            right = int(32767 * 0.55 * right_gain * value)
            frames.extend(struct.pack("<hh", left, right))

    def add_silence(duration):
        frames.extend(b"\x00\x00\x00\x00" * int(sample_rate * duration))

    add_silence(0.25)
    add_tone(0.8, 1.0, 0.0)
    add_silence(0.35)
    add_tone(0.8, 0.0, 1.0)
    add_silence(0.25)

    with wave.open(path, "wb") as handle:
        handle.setnchannels(2)
        handle.setsampwidth(2)
        handle.setframerate(sample_rate)
        handle.writeframes(bytes(frames))


def play_wav_file(path):
    if os.name == "nt":
        try:
            import winsound

            winsound.PlaySound(path, winsound.SND_FILENAME)
            return True, "played with winsound"
        except Exception as exc:
            return False, str(exc)

    players = [
        ["paplay", path],
        ["pw-play", path],
        ["aplay", path],
        ["ffplay", "-nodisp", "-autoexit", "-loglevel", "quiet", path],
    ]
    for command in players:
        if shutil.which(command[0]):
            rc, stdout, stderr = run_command_result(command, timeout=12)
            if rc == 0:
                return True, f"played with {command[0]}"
            return False, stderr or stdout or f"{command[0]} exited with code {rc}"
    return False, "No supported audio player found (paplay, pw-play, aplay or ffplay)"


def pc_auditor_sound_channel_test(non_interactive=False):
    fd, path = tempfile.mkstemp(prefix="auditor_sound_", suffix=".wav")
    os.close(fd)
    try:
        write_stereo_channel_wav(path)
        played, playback_details = play_wav_file(path)
        if not played:
            return status_item("Test sound channels", "UNKNOWN", playback_details)
        if non_interactive:
            return status_item("Test sound channels", "WARN", f"{playback_details}; no human confirmation")

        print("\nSound channel test")
        print("You should hear a tone on the LEFT channel, then a tone on the RIGHT channel.")
        status = ask_manual_status("Did you hear both left and right channels correctly?")
        note = input("Optional note (press Enter to skip): ").strip()
        if status == "OK":
            outcome = note or "User confirmed left/right sound output"
        elif status == "FAIL":
            outcome = note or "User reported sound output problem"
        else:
            outcome = note or "User skipped sound confirmation"
        return status_item("Test sound channels", status, f"{outcome} | {playback_details}")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def gui_available():
    if os.name == "nt":
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def ask_confirmation_result(test_name, ok_text, fail_text, skip_text, details=""):
    status = ask_manual_status(test_name)
    note = input("Optional note (press Enter to skip): ").strip()
    if status == "OK":
        result = note or ok_text
    elif status == "FAIL":
        result = note or fail_text
    else:
        result = note or skip_text
    if details:
        result += f" | {details}"
    return status_item(test_name, status, result)


def pc_auditor_lcd_panel_test(non_interactive=False):
    test_name = "Test LCD panel colors"
    if non_interactive:
        return status_item(test_name, "UNKNOWN", "Skipped; LCD color test needs technician confirmation")
    if not gui_available():
        return status_item(test_name, "UNKNOWN", "No graphical display available for LCD test")

    colors = [
        ("White", "#ffffff", "#000000"),
        ("Black", "#000000", "#ffffff"),
        ("Red", "#ff0000", "#ffffff"),
        ("Green", "#00ff00", "#000000"),
        ("Blue", "#0000ff", "#ffffff"),
        ("Cyan", "#00ffff", "#000000"),
        ("Magenta", "#ff00ff", "#ffffff"),
        ("Yellow", "#ffff00", "#000000"),
        ("Gray", "#808080", "#ffffff"),
    ]

    try:
        import tkinter as tk

        root = tk.Tk()
        root.title("LCD panel color test")
        root.attributes("-fullscreen", True)
        root.configure(bg=colors[0][1])
        index = {"value": 0}
        label = tk.Label(
            root,
            text="LCD COLOR TEST\nSpace/Right: next color\nLeft: previous color\nEsc: finish",
            font=("Arial", 28, "bold"),
            bg=colors[0][1],
            fg=colors[0][2],
            justify="center",
        )
        label.pack(expand=True, fill="both")

        def apply_color():
            name, bg, fg = colors[index["value"]]
            root.configure(bg=bg)
            label.configure(
                bg=bg,
                fg=fg,
                text=f"{name}\nCheck dead/lit pixels, lines, flicker and tint\nSpace/Right: next | Left: previous | Esc: finish",
            )

        def next_color(event=None):
            index["value"] = (index["value"] + 1) % len(colors)
            apply_color()

        def previous_color(event=None):
            index["value"] = (index["value"] - 1) % len(colors)
            apply_color()

        root.bind("<space>", next_color)
        root.bind("<Return>", next_color)
        root.bind("<Right>", next_color)
        root.bind("<Left>", previous_color)
        root.bind("<Escape>", lambda event=None: root.destroy())
        root.focus_force()
        apply_color()
        root.mainloop()
    except Exception as exc:
        return status_item(test_name, "UNKNOWN", f"Could not open LCD test window: {exc}")

    return ask_confirmation_result(
        test_name,
        "Technician confirmed LCD panel colors/pixels",
        "Technician reported LCD panel issue",
        "Technician skipped LCD panel confirmation",
    )


def pc_auditor_keyboard_feedback_test(non_interactive=False):
    test_name = "Test keyboard visual feedback"
    if non_interactive:
        return status_item(test_name, "UNKNOWN", "Skipped; keyboard test needs technician input")
    if not gui_available():
        return status_item(test_name, "UNKNOWN", "No graphical display available for keyboard test")

    details = {}
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title("Keyboard test")
        root.geometry("900x520")
        root.configure(bg="#101827")

        pressed = set()
        released = set()
        history = []

        title = tk.Label(
            root,
            text="KEYBOARD TEST",
            font=("Arial", 24, "bold"),
            bg="#101827",
            fg="#e5e7eb",
        )
        title.pack(pady=(20, 8))
        info = tk.Label(
            root,
            text="Press every key you want to verify. F12 finishes the window.",
            font=("Arial", 14),
            bg="#101827",
            fg="#cbd5e1",
        )
        info.pack(pady=(0, 16))
        count_label = tk.Label(root, font=("Consolas", 16), bg="#101827", fg="#93c5fd")
        count_label.pack(pady=8)
        history_box = tk.Text(root, height=14, bg="#020617", fg="#e5e7eb", insertbackground="#e5e7eb")
        history_box.pack(fill="both", expand=True, padx=24, pady=(8, 24))
        history_box.insert("1.0", "Key log will appear here...\n")
        history_box.configure(state="disabled")

        def refresh():
            count_label.configure(text=f"Pressed: {len(pressed)} unique | Released: {len(released)} unique")
            history_box.configure(state="normal")
            history_box.delete("1.0", tk.END)
            history_box.insert("1.0", "\n".join(history[-40:]) or "No keys pressed yet")
            history_box.configure(state="disabled")

        def key_name(event):
            return event.keysym or event.char or "Unknown"

        def on_press(event):
            name = key_name(event)
            if name == "F12":
                root.destroy()
                return
            pressed.add(name)
            history.append(f"DOWN  {name}")
            refresh()

        def on_release(event):
            name = key_name(event)
            released.add(name)
            history.append(f"UP    {name}")
            refresh()

        root.bind("<KeyPress>", on_press)
        root.bind("<KeyRelease>", on_release)
        root.focus_force()
        refresh()
        root.mainloop()
        details = {"pressed": len(pressed), "released": len(released)}
    except Exception as exc:
        return status_item(test_name, "UNKNOWN", f"Could not open keyboard test window: {exc}")

    return ask_confirmation_result(
        test_name,
        "Technician confirmed tested keys",
        "Technician reported keyboard issue",
        "Technician skipped keyboard confirmation",
        details=f"{details.get('pressed', 0)} unique key press(es), {details.get('released', 0)} release event(s)",
    )


def pc_auditor_pointer_canvas_test(test_name, instructions, non_interactive=False):
    if non_interactive:
        return status_item(test_name, "UNKNOWN", "Skipped; pointer/touch test needs technician input")
    if not gui_available():
        return status_item(test_name, "UNKNOWN", "No graphical display available for pointer/touch test")

    stats = {"motion": 0, "clicks": 0}
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title(test_name)
        root.geometry("900x600")
        root.configure(bg="#101827")

        label = tk.Label(
            root,
            text=f"{test_name}\n{instructions}\nMove/draw/click across the area. Esc finishes.",
            font=("Arial", 14),
            bg="#101827",
            fg="#e5e7eb",
            justify="center",
        )
        label.pack(fill="x", pady=(14, 10))
        canvas = tk.Canvas(root, bg="#ffffff", highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        last = {"x": None, "y": None}

        def mark(event, click=False):
            if click:
                stats["clicks"] += 1
                canvas.create_oval(event.x - 8, event.y - 8, event.x + 8, event.y + 8, fill="#2563eb", outline="")
            else:
                stats["motion"] += 1
            if last["x"] is not None and last["y"] is not None:
                canvas.create_line(last["x"], last["y"], event.x, event.y, fill="#dc2626", width=3)
            last["x"], last["y"] = event.x, event.y

        canvas.bind("<Motion>", lambda event: mark(event))
        canvas.bind("<B1-Motion>", lambda event: mark(event))
        canvas.bind("<ButtonPress-1>", lambda event: mark(event, click=True))
        canvas.bind("<ButtonPress-2>", lambda event: mark(event, click=True))
        canvas.bind("<ButtonPress-3>", lambda event: mark(event, click=True))
        root.bind("<Escape>", lambda event=None: root.destroy())
        root.mainloop()
    except Exception as exc:
        return status_item(test_name, "UNKNOWN", f"Could not open pointer/touch test window: {exc}")

    return ask_confirmation_result(
        test_name,
        "Technician confirmed pointer/touch movement and buttons",
        "Technician reported pointer/touch problem",
        "Technician skipped pointer/touch confirmation",
        details=f"{stats['motion']} motion event(s), {stats['clicks']} click/touch event(s)",
    )


def webcam_device_detected():
    system = platform.system().lower()
    if system.startswith("win"):
        camera = first_line(
            powershell_value(
                "Get-CimInstance Win32_PnPEntity | "
                "Where-Object { $_.Name -match 'Camera|Webcam' } | Select-Object -First 1 -ExpandProperty Name"
            )
        )
        return camera
    if system == "linux" and os.path.isdir("/dev"):
        devices = [name for name in os.listdir("/dev") if name.startswith("video")]
        return ", ".join(devices)
    return ""


def pc_auditor_webcam_preview_test(non_interactive=False):
    test_name = "Test webcam preview"
    detected = webcam_device_detected()
    if not detected:
        return status_item(test_name, "UNKNOWN", "No webcam device detected")
    if non_interactive:
        return status_item(test_name, "WARN", f"Webcam detected but not previewed: {detected}")

    preview_details = f"Detected: {detected}"
    try:
        import cv2

        cap = cv2.VideoCapture(0)
        if cap.isOpened():
            print("\nWebcam preview")
            print("A webcam window will open. Press q or Esc to close it.")
            frame_count = 0
            start = time.time()
            while time.time() - start < 10:
                ok, frame = cap.read()
                if not ok:
                    break
                frame_count += 1
                cv2.imshow("Webcam preview - press q or Esc", frame)
                key = cv2.waitKey(1) & 0xFF
                if key in {ord("q"), 27}:
                    break
            cap.release()
            cv2.destroyAllWindows()
            preview_details += f"; captured {frame_count} frame(s)"
        else:
            preview_details += "; OpenCV could not open camera"
    except Exception as exc:
        preview_details += f"; preview dependency unavailable ({exc})"
        print("\nWebcam preview")
        print("Open the OS Camera/Webcam app now and verify the image manually.")

    return ask_confirmation_result(
        test_name,
        "Technician confirmed webcam video",
        "Technician reported webcam/video quality issue",
        "Technician skipped webcam confirmation",
        details=preview_details,
    )


def pc_auditor_memory_smoke_test(memory_mb=128):
    test_name = "Test memory smoke"
    try:
        memory_mb = int(memory_mb)
    except (TypeError, ValueError):
        memory_mb = 128
    if memory_mb <= 0:
        return status_item(test_name, "UNKNOWN", "Skipped; memory test size is 0 MB")
    memory_mb = max(8, min(memory_mb, 2048))

    try:
        chunks = []
        for index in range(memory_mb):
            value = index % 251
            chunk = bytearray(1024 * 1024)
            for offset in range(0, len(chunk), 4096):
                chunk[offset] = value
            chunks.append((value, chunk))
        for value, chunk in chunks:
            for offset in range(0, len(chunk), 4096):
                if chunk[offset] != value:
                    return status_item(test_name, "FAIL", f"Memory verify mismatch at offset {offset}")
        return status_item(
            test_name,
            "OK",
            f"Allocated and verified {memory_mb} MB inside the running OS; not a full boot-level memtest",
        )
    except MemoryError:
        return status_item(test_name, "FAIL", f"Could not allocate {memory_mb} MB for memory smoke test")
    except Exception as exc:
        return status_item(test_name, "FAIL", f"Memory smoke test failed: {exc}")


def usb_device_listing():
    system = platform.system().lower()
    if system.startswith("win"):
        output = powershell_value(
            "Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPDeviceID -like 'USB*' } | "
            "Select-Object -First 80 -ExpandProperty Name"
        )
        lines = [line.strip() for line in output.splitlines() if line.strip()]
        return lines
    if system == "linux":
        output = run_command(["lsusb"], timeout=6)
        lines = [line.strip() for line in output.splitlines() if line.strip()]
        if lines:
            return lines
        if os.path.isdir("/sys/bus/usb/devices"):
            return [name for name in os.listdir("/sys/bus/usb/devices") if ":" not in name]
    return []


def pc_auditor_usb_detection_test(expected_usb_devices=-1, non_interactive=False):
    test_name = "Test USB device detection"
    lines = usb_device_listing()
    expected = expected_usb_devices
    try:
        expected = int(expected)
    except (TypeError, ValueError):
        expected = -1

    if expected < 0 and not non_interactive:
        print("\nUSB detection test")
        print("Insert known-good USB devices into the ports you want to verify.")
        answer = input("Expected number of visible USB test devices [Enter=do not compare]: ").strip()
        if answer:
            try:
                expected = int(answer)
            except ValueError:
                expected = -1

    count = len(lines)
    details = f"{count} USB device/controller item(s) visible"
    if lines:
        details += " | " + "; ".join(lines[:8])

    if expected >= 0:
        details += f" | expected at least {expected}"
        status = "OK" if count >= expected else "FAIL"
    else:
        status = "OK" if count else "WARN"
    return status_item(test_name, status, details)


def pc_auditor_battery_status_test():
    test_name = "Test notebook battery"
    system = platform.system().lower()
    if system.startswith("win"):
        output = powershell_value(
            "Get-CimInstance Win32_Battery | Select-Object -First 4 | "
            "ForEach-Object { $_.Name + '|charge=' + $_.EstimatedChargeRemaining + '%|status=' + "
            "$_.BatteryStatus + '|runtime=' + $_.EstimatedRunTime }"
        )
        lines = [line.strip() for line in output.splitlines() if line.strip()]
        if lines:
            return status_item("Test notebook battery", "OK", "; ".join(lines))
        return status_item(test_name, "UNKNOWN", "No notebook battery reported by Windows")

    if system == "linux":
        batteries = []
        base = "/sys/class/power_supply"
        if os.path.isdir(base):
            for name in os.listdir(base):
                if not name.startswith("BAT"):
                    continue
                path = os.path.join(base, name)
                capacity = read_text_file(os.path.join(path, "capacity"))
                status = read_text_file(os.path.join(path, "status"))
                full = read_text_file(os.path.join(path, "energy_full")) or read_text_file(os.path.join(path, "charge_full"))
                design = read_text_file(os.path.join(path, "energy_full_design")) or read_text_file(os.path.join(path, "charge_full_design"))
                health = ""
                try:
                    if full and design:
                        health = f" health={round((int(full) / int(design)) * 100)}%"
                except (ValueError, ZeroDivisionError):
                    health = ""
                batteries.append(f"{name}: charge={capacity}% status={status}{health}".strip())
        if batteries:
            return status_item(test_name, "OK", "; ".join(batteries))
        return status_item(test_name, "UNKNOWN", "No notebook battery reported by Linux")

    return status_item(test_name, "UNKNOWN", f"Unsupported OS: {platform.system()}")


def collect_auditor_style_tests(args):
    print("\nPC Auditor style tests")
    print("These checks are safe: they only read devices/media, play a short sound, and test server packets.\n")

    items = [
        pc_auditor_disk_health_test(),
        pc_auditor_memory_smoke_test(args.memory_test_mb),
        pc_auditor_battery_status_test(),
        pc_auditor_server_packet_test(args.server, attempts=args.test_attempts, timeout=args.timeout),
        pc_auditor_usb_detection_test(args.expected_usb_devices, non_interactive=args.non_interactive_tests),
        pc_auditor_optical_read_test(args),
        pc_auditor_removable_read_test(args),
        pc_auditor_sound_channel_test(non_interactive=args.non_interactive_tests),
        pc_auditor_lcd_panel_test(non_interactive=args.non_interactive_tests),
        pc_auditor_keyboard_feedback_test(non_interactive=args.non_interactive_tests),
        pc_auditor_pointer_canvas_test(
            "Test touchscreen",
            "If the machine has a touchscreen, draw lines with your finger across the panel.",
            non_interactive=args.non_interactive_tests,
        ),
        pc_auditor_webcam_preview_test(non_interactive=args.non_interactive_tests),
        pc_auditor_pointer_canvas_test(
            "Test touchpad",
            "Use the integrated touchpad for movement, left/right click and button checks.",
            non_interactive=args.non_interactive_tests,
        ),
    ]
    return {"items": items, "summary": summarize_status(items)}


def collect_windows_device_tests(args):
    print("\nRequired Windows device tests")
    print("USB detection/count, optical read, USB/SD read and left/right sound.\n")

    if not platform.system().lower().startswith("win"):
        return {
            "items": [
                status_item(
                    "Windows device tests",
                    "UNKNOWN",
                    f"Windows-only test suite was requested on {platform.system()}",
                )
            ],
            "summary": {"UNKNOWN": 1},
        }

    items = [
        pc_auditor_usb_detection_test(args.expected_usb_devices, non_interactive=args.non_interactive_tests),
        pc_auditor_optical_read_test(args),
        pc_auditor_removable_read_test(args),
        pc_auditor_sound_channel_test(non_interactive=args.non_interactive_tests),
    ]
    return {"items": items, "summary": summarize_status(items)}


MANUAL_PORT_GROUPS = [
    (
        "USB-A",
        "Plug a known-good USB mouse or flash drive into this USB-A port. Confirm movement/click or file access.",
    ),
    (
        "USB-C / Thunderbolt",
        "Plug a known-good USB-C device/dock. Confirm data access. If supported, also confirm charging or display output.",
    ),
    (
        "LAN / Ethernet",
        "Connect a known-good Ethernet cable. Confirm link lights and network access.",
    ),
    (
        "HDMI",
        "Connect a known-good external monitor/TV. Confirm image output.",
    ),
    (
        "DisplayPort / Mini DisplayPort",
        "Connect a known-good external monitor. Confirm image output.",
    ),
    (
        "VGA",
        "Connect a known-good VGA monitor/projector. Confirm image output.",
    ),
    (
        "Audio jack",
        "Connect known-good headphones/headset. Confirm left/right audio and microphone if combo jack.",
    ),
    (
        "SD / microSD reader",
        "Insert a known-good SD/microSD card. Confirm it appears and files can be read.",
    ),
    (
        "Charging / power port",
        "Connect a known-good charger. Confirm charging indicator and OS charging status.",
    ),
    (
        "Docking port",
        "Connect a compatible dock. Confirm dock USB/network/display functions if available.",
    ),
]

MANUAL_DEVICE_TESTS = [
    (
        "Manual display panel",
        "Check LCD/OLED panel for image, brightness control, dead lines, flicker, color tint and backlight issues.",
    ),
    (
        "Manual touchscreen",
        "If present, touch multiple screen areas and confirm accurate response.",
    ),
    (
        "Manual keyboard",
        "Type letters, numbers, space, enter, backspace and modifier keys. Confirm all tested keys work.",
    ),
    (
        "Manual touchpad/buttons",
        "Move pointer, left/right click, tap, scroll and gesture if supported.",
    ),
    (
        "Manual camera",
        "Open camera app/test tool and confirm video image appears.",
    ),
    (
        "Manual microphone",
        "Record or use an audio test and confirm microphone input.",
    ),
    (
        "Manual speakers",
        "Play a known-good audio sample and confirm speakers work without distortion.",
    ),
    (
        "Manual Wi-Fi",
        "Connect to a known-good Wi-Fi network and confirm internet/LAN access.",
    ),
    (
        "Manual Bluetooth",
        "Pair or detect a known-good Bluetooth device.",
    ),
    (
        "Manual optical drive",
        "Insert a known-good finalized CD/DVD. Confirm the drive reads files/folders.",
    ),
    (
        "Manual battery",
        "Unplug charger briefly and confirm battery powers the machine. Reconnect charger and confirm charging.",
    ),
    (
        "Manual fingerprint reader",
        "If present, open OS sign-in/security test and confirm the reader responds.",
    ),
    (
        "Manual function keys",
        "Test brightness, volume, mute, Wi-Fi toggle and other function keys if present.",
    ),
    (
        "Manual sleep/wake",
        "Put the device to sleep and wake it. Confirm it resumes normally.",
    ),
]


def ask_manual_status(prompt):
    while True:
        answer = input(f"{prompt} [p=pass / f=fail / s=skip]: ").strip().lower()
        if answer in {"p", "pass", "y", "yes"}:
            return "OK"
        if answer in {"f", "fail", "n", "no"}:
            return "FAIL"
        if answer in {"s", "skip", ""}:
            return "UNKNOWN"
        print("Please enter p, f, or s.")


def ask_manual_count(prompt):
    while True:
        answer = input(f"{prompt} [0 to skip]: ").strip()
        if answer == "":
            return 0
        try:
            count = int(answer)
        except ValueError:
            print("Please enter a number.")
            continue
        if count < 0:
            print("Please enter 0 or a positive number.")
            continue
        return count


def manual_details_for_status(status, note, instruction):
    if status == "OK":
        result = note or "Manual test passed"
    elif status == "FAIL":
        result = note or "Manual test failed"
    else:
        result = note or "Manual test skipped or device/port not present"
    return f"{result} | Procedure: {instruction}"


def append_manual_test(items, name, instruction):
    print(instruction)
    status = ask_manual_status(name)
    note = input("Optional note (press Enter to skip): ").strip()
    items.append(status_item(name, status, manual_details_for_status(status, note, instruction)))
    print("")


def collect_manual_diagnostics():
    print("\nFull manual device and port checks")
    print("Use known-good test devices. Skipped or non-present devices are saved as UNKNOWN.")
    print("For ports, enter how many physical ports you want to test on this machine.\n")

    items = []

    print("Physical ports")
    for port_name, instruction in MANUAL_PORT_GROUPS:
        count = ask_manual_count(f"How many {port_name} port(s) will you test?")
        for index in range(1, count + 1):
            append_manual_test(items, f"Manual {port_name} port {index}", instruction)

    print("Built-in devices and functional checks")
    for name, instruction in MANUAL_DEVICE_TESTS:
        append_manual_test(items, name, instruction)

    return {"items": items, "summary": summarize_status(items)}


def merge_diagnostics(auto_diagnostics, manual_diagnostics):
    items = []
    if isinstance(auto_diagnostics, dict):
        items.extend(auto_diagnostics.get("items", []))
    if isinstance(manual_diagnostics, dict):
        items.extend(manual_diagnostics.get("items", []))
    return {"items": items, "summary": summarize_status(items)}


def collect_inventory():
    system = platform.system().lower()

    if system.startswith("win"):
        chassis = first_line(powershell_value("(Get-CimInstance Win32_SystemEnclosure).ChassisTypes[0]"))
        product_type = chassis_label(chassis)
        disk_info = first_line(
            powershell_value(
                "Get-PhysicalDisk | Select-Object -First 1 | "
                "ForEach-Object { ($_.MediaType.ToString()) + '|' + ([math]::Round($_.Size / 1GB).ToString() + ' GB') }"
            )
        )
        disk_type, disk_size = "", ""
        if "|" in disk_info:
            disk_type, disk_size = [part.strip() for part in disk_info.split("|", 1)]
            if disk_type.upper() == "SSD":
                disk_type = "SATA SSD"

        inventory = {
            "Product type": product_type,
            "Manufacturer": first_line(powershell_value("(Get-CimInstance Win32_ComputerSystem).Manufacturer")),
            "Model": first_line(powershell_value("(Get-CimInstance Win32_ComputerSystem).Model")),
            "Chassis type": product_type,
            "Part number": first_line(powershell_value("(Get-CimInstance Win32_ComputerSystemProduct).SKUNumber")),
            "Serial number": first_line(powershell_value("(Get-CimInstance Win32_BIOS).SerialNumber")),
            "Processor": first_line(powershell_value("(Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)")),
            "RAM size": first_line(
                powershell_value("[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB).ToString() + ' GB'")
            ),
            "Disk type": disk_type,
            "Disk size": disk_size,
            "Optical drive": first_line(powershell_value("(Get-CimInstance Win32_CDROMDrive | Select-Object -First 1 -ExpandProperty Name)")),
            "Video model": first_line(powershell_value("(Get-CimInstance Win32_VideoController | Select-Object -First 1 -ExpandProperty Name)")),
            "LCD type": infer_screen_type_from_text(windows_monitor_text()),
        }
    elif system == "linux":
        disk_type, disk_size = linux_disk_info()
        cpu = ""
        for line in read_text_file("/proc/cpuinfo").splitlines():
            if line.lower().startswith("model name"):
                cpu = line.split(":", 1)[1].strip()
                break
        ram = ""
        for line in read_text_file("/proc/meminfo").splitlines():
            if line.startswith("MemTotal:"):
                ram = format_gb(int(line.split()[1]) * 1024)
                break
        video = ""
        for line in run_command(["lspci"]).splitlines():
            lower = line.lower()
            if "vga compatible controller" in lower or "3d controller" in lower or "display controller" in lower:
                video = line.split(":", 2)[-1].strip()
                break
        product_type = chassis_label(read_text_file("/sys/class/dmi/id/chassis_type"))
        inventory = {
            "Product type": product_type,
            "Manufacturer": read_text_file("/sys/class/dmi/id/sys_vendor"),
            "Model": read_text_file("/sys/class/dmi/id/product_name"),
            "Chassis type": product_type,
            "Part number": read_text_file("/sys/class/dmi/id/product_sku"),
            "Serial number": read_text_file("/sys/class/dmi/id/product_serial") or read_text_file("/sys/class/dmi/id/board_serial"),
            "Processor": cpu,
            "RAM size": ram,
            "Disk type": disk_type,
            "Disk size": disk_size,
            "Video model": video,
            "LCD type": infer_screen_type_from_text(linux_monitor_text()),
        }
    else:
        inventory = {}

    return {key: value for key, value in inventory.items() if value and str(value).strip()}


def post_json(server_url, path, payload, timeout=8):
    url = server_url.rstrip("/") + path
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urlrequest.Request(
        url,
        data=body,
        method="POST",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
        },
    )

    try:
        with urlrequest.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urlerror.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw)
            message = data.get("error", raw)
        except json.JSONDecodeError:
            message = raw or str(exc)
        raise RuntimeError(message) from exc
    except urlerror.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    except TimeoutError as exc:
        raise RuntimeError("request timed out") from exc


def get_json(server_url, path, timeout=8):
    url = server_url.rstrip("/") + path
    req = urlrequest.Request(url, method="GET", headers={"Accept": "application/json"})

    try:
        with urlrequest.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urlerror.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw)
            message = data.get("error", raw)
        except json.JSONDecodeError:
            message = raw or str(exc)
        raise RuntimeError(message) from exc
    except urlerror.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    except TimeoutError as exc:
        raise RuntimeError("request timed out") from exc


def manual_tests_required(server_url, timeout=8):
    result = get_json(server_url, "/settings", timeout=timeout)
    settings = result.get("settings", {})
    return bool(settings.get("manual_tests_required", False))


def register_client(args, include_manual_tests=False, include_auditor_tests=False, include_windows_tests=False):
    inventory = {} if args.no_inventory else collect_inventory()
    diagnostics = None
    if not args.no_diagnostics:
        diagnostics = collect_device_diagnostics()
    if include_auditor_tests:
        auditor_diagnostics = collect_auditor_style_tests(args)
        diagnostics = merge_diagnostics(diagnostics, auditor_diagnostics)
    if include_windows_tests:
        windows_diagnostics = collect_windows_device_tests(args)
        diagnostics = merge_diagnostics(diagnostics, windows_diagnostics)
    if include_manual_tests:
        manual_diagnostics = collect_manual_diagnostics()
        diagnostics = merge_diagnostics(diagnostics, manual_diagnostics)
    if diagnostics:
        inventory["Diagnostics"] = diagnostics

    payload = {
        "name": args.name,
        "username": args.username,
        "role": args.role,
        "app_version": args.app_version,
        "platform": platform_label(),
        "inventory": inventory,
    }
    return post_json(args.server, "/clients/register", payload, timeout=args.timeout)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Register this workstation with the Enterprise Desktop System network job server."
    )
    parser.add_argument("--server", default=DEFAULT_SERVER, help="Job server URL, e.g. http://192.168.1.50:8765")
    parser.add_argument("--name", default=DEFAULT_NAME, help="Client/workstation name shown in Available Clients")
    parser.add_argument("--username", default=getpass.getuser(), help="Username shown for this client")
    parser.add_argument("--role", default=DEFAULT_ROLE, help="Client role label")
    parser.add_argument("--app-version", default="register-client-script", help="Version label shown in the server")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL, help="Heartbeat interval in seconds")
    parser.add_argument("--timeout", type=int, default=8, help="HTTP timeout in seconds")
    parser.add_argument("--once", action="store_true", help="Register once and exit")
    parser.add_argument("--no-inventory", action="store_true", help="Register without sending hardware details")
    parser.add_argument("--no-diagnostics", action="store_true", help="Register without sending device diagnostics")
    parser.add_argument(
        "--auditor-tests",
        action="store_true",
        help="Run full Aiken-like PC-Auditor-style tests for disks, ports, sound, LCD, keyboard, camera and input devices",
    )
    parser.add_argument(
        "--aiken-tests",
        action="store_true",
        help="Alias for --auditor-tests with the full Aiken-like device test suite",
    )
    parser.add_argument(
        "--windows-device-tests",
        action="store_true",
        help="Run the required Windows tests: USB count, optical read, USB/SD read and left/right sound",
    )
    parser.add_argument(
        "--non-interactive-tests",
        action="store_true",
        help="Do not prompt during auditor tests; skip tests that need technician confirmation",
    )
    parser.add_argument("--optical-test-path", default="", help="Drive/path for optical media test, e.g. D: or /media/user/DVD")
    parser.add_argument("--usb-test-path", default="", help="Drive/path for removable USB/SD media test, e.g. E: or /media/user/USB")
    parser.add_argument(
        "--expected-usb-devices",
        type=int,
        default=-1,
        help="Expected visible USB device count for the USB detection comparison; -1 disables comparison",
    )
    parser.add_argument(
        "--memory-test-mb",
        type=int,
        default=128,
        help="Memory size in MB for the safe in-OS memory smoke test",
    )
    parser.add_argument("--test-attempts", type=int, default=3, help="Server packet attempts for the network test")
    parser.add_argument("--manual-tests", action="store_true", help="Run interactive manual device checks and send the results")
    parser.add_argument(
        "--ignore-required-manual-tests",
        action="store_true",
        help="Do not auto-run manual tests even if the server marks them as required",
    )
    parser.add_argument("--print-diagnostics", action="store_true", help="Print local device diagnostics and exit")
    return parser.parse_args()


def main():
    args = parse_args()
    args.server = args.server.rstrip("/")

    if not args.server.startswith(("http://", "https://")):
        args.server = "http://" + args.server

    if args.interval < 5:
        print("ERROR: interval must be at least 5 seconds", file=sys.stderr)
        return 2

    if args.print_diagnostics:
        diagnostics = collect_device_diagnostics()
        if args.auditor_tests or args.aiken_tests:
            diagnostics = merge_diagnostics(diagnostics, collect_auditor_style_tests(args))
        if args.windows_device_tests:
            diagnostics = merge_diagnostics(diagnostics, collect_windows_device_tests(args))
        if args.manual_tests:
            diagnostics = merge_diagnostics(diagnostics, collect_manual_diagnostics())
        print(json.dumps(diagnostics, indent=2, ensure_ascii=False))
        return 0

    manual_tests_done_this_session = False
    auditor_tests_done_this_session = False
    windows_tests_done_this_session = False

    while True:
        try:
            required_by_server = False
            if not args.ignore_required_manual_tests:
                try:
                    required_by_server = manual_tests_required(args.server, timeout=args.timeout)
                except RuntimeError as exc:
                    print(f"could not read server requirement: {exc}", file=sys.stderr, flush=True)

            include_manual_tests = False
            if args.manual_tests and not manual_tests_done_this_session:
                include_manual_tests = True
            elif required_by_server and not manual_tests_done_this_session:
                print("Server requires manual tests. Starting manual checklist now.", flush=True)
                include_manual_tests = True

            include_auditor_tests = (args.auditor_tests or args.aiken_tests) and not auditor_tests_done_this_session
            include_windows_tests = args.windows_device_tests and not windows_tests_done_this_session
            result = register_client(
                args,
                include_manual_tests=include_manual_tests,
                include_auditor_tests=include_auditor_tests,
                include_windows_tests=include_windows_tests,
            )
            if include_auditor_tests:
                auditor_tests_done_this_session = True
            if include_windows_tests:
                windows_tests_done_this_session = True
            if include_manual_tests:
                manual_tests_done_this_session = True
            last_seen = result.get("last_seen", "unknown")
            print(f"registered {args.name} at {last_seen} via {args.server}", flush=True)
        except RuntimeError as exc:
            print(f"register failed: {exc}", file=sys.stderr, flush=True)
            if args.once:
                return 1

        if args.once:
            return 0

        time.sleep(args.interval)


if __name__ == "__main__":
    raise SystemExit(main())
