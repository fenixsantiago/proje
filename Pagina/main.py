import os
import sys
import streamlit.web.cli as stcli

if __name__ == '__main__':
    # Validar si el programa está corriendo empaquetado en el EXE o en desarrollo
    if hasattr(sys, '_MEIPASS'):
        # Ruta dentro de la carpeta temporal del .EXE
        base_path = sys._MEIPASS
    else:
        # Ruta en tu carpeta de desarrollo habitual
        base_path = os.path.dirname(__file__)

    script_path = os.path.join(base_path, 'index.py')
    
    # Parámetros para ejecutar Streamlit internamente
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())