import os
import runpy


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
runpy.run_path(os.path.join(BASE_DIR, "enterprise_tabs_app.pyw"), run_name="__main__")
