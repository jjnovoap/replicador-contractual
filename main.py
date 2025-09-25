import os
import sys

def menu():
    print("Seleccione el entorno que desea ejecutar:")
    print("1. SolicitudOrden")
    print("2. SolicitudOCD_OSU")
    print("3. SolicitudEA")
    return input("Ingrese el número de la opción: ")

def ejecutar(opcion):
    entornos = {
        "1": "SolicitudOrden",
        "2": "SolicitudOCD_OSU",
        "3": "SolicitudEA"
    }
    if opcion in entornos:
        entorno = entornos[opcion]
        script = os.path.join(entorno, "main.py")
        os.system(f"{sys.executable} {script}")
    else:
        print("Opción inválida")

if __name__ == "__main__":
    ejecutar(menu())