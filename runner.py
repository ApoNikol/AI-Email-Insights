import os, sys
import streamlit.web.cli as stcli

def resolve_path(path):
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(bundle_dir, path)

if __name__ == "__main__":
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("ui.py")
        "--global.developmentMode=false",
    ]

    sys.exit(stcli.main())
