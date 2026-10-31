"""Crea el ejecutable portable desde Windows: python compilar_windows.py."""

import subprocess
import sys
from pathlib import Path


def main():
    if sys.platform != "win32":
        print("El .exe debe compilarse en Windows. Ejecuta CREAR_EXE_WINDOWS.bat en ese equipo.")
        return 1

    carpeta = Path(__file__).resolve().parent
    recursos = (carpeta / "FONDO.png", carpeta / "assets" / "bancolombia.png")
    for recurso in recursos:
        if not recurso.is_file():
            print(f"Falta el recurso: {recurso}")
            return 1

    comando = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", "Estructurador",
        "--distpath", str(carpeta / "dist"),
        "--workpath", str(carpeta / "build"),
        "--specpath", str(carpeta / "build"),
        "--icon", str(recursos[1]),
        "--add-data", f"{recursos[0]}:.",
        "--add-data", f"{recursos[1]}:assets",
        str(carpeta / "app.py"),
    ]
    resultado = subprocess.run(comando, cwd=carpeta, check=False)
    if resultado.returncode:
        return resultado.returncode
    print(f"Ejecutable creado: {carpeta / 'dist' / 'Estructurador.exe'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
