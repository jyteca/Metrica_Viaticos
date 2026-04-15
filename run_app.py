import streamlit.web.cli as stcli
import sys
import os

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        # Ejecutando como archivo compilado
        bundle_dir = sys._MEIPASS
    else:
        # Ejecutando como script normal
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    app_path = os.path.join(bundle_dir, "app.py")
    
    # Ocultar terminal inicial o desarrollo de Streamlit
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENTMODE"] = "false"
    
    # Iniciar la aplicacion de Streamlit
    sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())
