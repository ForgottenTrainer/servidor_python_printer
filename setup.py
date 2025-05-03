import sys
import subprocess
import platform
import os

def main():
    print("==== INSTALACIÓN DE DEPENDENCIAS PARA SERVIDOR DE IMPRESIÓN ====\n")
    
    # Verificar que es Windows
    if platform.system() != "Windows":
        print("❌ Este servidor solo es compatible con Windows.")
        return
    
    # Lista de dependencias básicas
    dependencies = [
        "flask",
        "werkzeug",
        "pywin32",
        "reportlab",
        "pillow"
    ]
    
    # Instalar cada dependencia
    print("Instalando dependencias...")
    for dependency in dependencies:
        print(f"Instalando {dependency}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", dependency])
            print(f"✅ {dependency} instalado correctamente")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error al instalar {dependency}: {e}")
    
    print("\n==== INSTALACIÓN COMPLETADA ====")
    print("\nAhora puedes ejecutar el servidor con:")
    print("python app.py")

if __name__ == "__main__":

import sys
import subprocess
import platform
import os

def main():
    print("==== INSTALACIÓN DE DEPENDENCIAS PARA SERVIDOR DE IMPRESIÓN ====\n")
    
    # Verificar que es Windows
    if platform.system() != "Windows":
        print("❌ Este servidor solo es compatible con Windows.")
        return
    
    # Lista de dependencias básicas
    dependencies = [
        "flask",
        "werkzeug",
        "pywin32",
        "reportlab",
        "pillow"
    ]
    
    # Instalar cada dependencia
    print("Instalando dependencias...")
    for dependency in dependencies:
        print(f"Instalando {dependency}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", dependency])
            print(f"✅ {dependency} instalado correctamente")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error al instalar {dependency}: {e}")
    
    print("\n==== INSTALACIÓN COMPLETADA ====")
    print("\nAhora puedes ejecutar el servidor con:")
    print("python app.py")

if __name__ == "__main__":

    main()