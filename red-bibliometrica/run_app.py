#!/usr/bin/env python3
import sys
import subprocess
import socket
import webbrowser
import threading
import time

def install_dependencies():
    """
    Checks for the presence of needed libraries and installs them using pip if missing.
    """
    required_packages = {
        'flask': 'flask',
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'numpy': 'numpy'
    }
    
    installed_any = False
    for module_name, pip_name in required_packages.items():
        try:
            __import__(module_name)
        except ImportError:
            print(f"[*] Instalando librería de Python requerida: '{pip_name}'...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
                print(f"[+] '{pip_name}' se instaló correctamente.")
                installed_any = True
            except subprocess.CalledProcessError as e:
                print(f"[!] Error al instalar '{pip_name}': {e}")
                print("[!] Intenta instalarlo manualmente ejecutando: pip install " + pip_name)
                sys.exit(1)
                
    if installed_any:
        print("[+] Todas las dependencias preparadas con éxito.\n")

def find_free_port(start_port=5000):
    """
    Finds a free port on localhost starting from start_port.
    Avoids conflict if port 5000 is occupied (e.g. macOS AirPlay receiver).
    """
    port = start_port
    while port < 6000:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            # If connect_ex returns anything other than 0, the port is free
            if s.connect_ex(('localhost', port)) != 0:
                return port
        port += 1
    return start_port

def main():
    print("=" * 60)
    print("     INICIANDO ANALIZADOR DE REDES BIBLIOMÉTRICAS     ")
    print("=" * 60)
    
    # 1. Self-install missing libraries
    install_dependencies()
    
    # 2. Find a free port
    port = find_free_port(5000)
    url = f"http://localhost:{port}"
    
    print(f"[*] Levantando servidor local en: {url}")
    print("[*] Abriendo aplicación en tu navegador web de forma automática...")
    
    # 3. Import Flask app now that we've checked dependencies
    try:
        from app import app
    except Exception as e:
        print(f"[!] Error al cargar el backend: {e}")
        sys.exit(1)
        
    # 4. Open default web browser after server starts up
    def open_browser():
        time.sleep(1.2)
        webbrowser.open(url)
        
    threading.Thread(target=open_browser, daemon=True).start()
    
    # 5. Run Flask server locally
    try:
        # Disable Flask debug mode output and run quietly
        import logging
        log = logging.getLogger('werkzeug')
        log.setLevel(logging.ERROR)
        
        print("\n[✔] Aplicación lista y corriendo. Para cerrarla presiona Ctrl+C en esta ventana.")
        print("-" * 60)
        app.run(host='localhost', port=port, debug=False)
    except KeyboardInterrupt:
        print("\n\n[!] Servidor detenido por el usuario. ¡Adiós!")
    except Exception as e:
        print(f"\n[!] Error al iniciar el servidor Flask: {e}")

if __name__ == '__main__':
    main()
