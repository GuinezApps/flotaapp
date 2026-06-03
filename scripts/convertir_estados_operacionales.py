import os
import sys
import django
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.append(str(BASE_DIR))

os.environ.setdefault(
    "DJANGO_SETTINGS_MODULE",
    "config.settings"
)

django.setup()


from flota.models import Vehiculo


conversiones = {
    "Operativo": {
        "estado_operacional": "Operativo",
        "subestado_no_operativo": None,
    },
    "En Mantencion": {
        "estado_operacional": "No operativo",
        "subestado_no_operativo": "Mantencion",
    },
    "FueraServicio": {
        "estado_operacional": "No operativo",
        "subestado_no_operativo": "Detenido",
    },
    "No informado": {
        "estado_operacional": "No informado",
        "subestado_no_operativo": None,
    },
}


actualizados = 0
sin_cambios = 0
desconocidos = []


for vehiculo in Vehiculo.objects.all():

    estado_actual = vehiculo.estado_operacional

    if estado_actual not in conversiones:
        desconocidos.append(
            f"{vehiculo.patente}: {estado_actual}"
        )
        continue

    nuevo_estado = conversiones[estado_actual]["estado_operacional"]
    nuevo_subestado = conversiones[estado_actual]["subestado_no_operativo"]

    if (
        vehiculo.estado_operacional == nuevo_estado
        and vehiculo.subestado_no_operativo == nuevo_subestado
    ):
        sin_cambios += 1
        continue

    vehiculo.estado_operacional = nuevo_estado
    vehiculo.subestado_no_operativo = nuevo_subestado
    vehiculo.save()

    actualizados += 1


print("")
print("=== CONVERSIÓN DE ESTADOS OPERACIONALES ===")
print("")
print(f"Actualizados: {actualizados}")
print(f"Sin cambios: {sin_cambios}")
print(f"Estados desconocidos: {len(desconocidos)}")

if desconocidos:
    print("")
    print("Vehículos con estado desconocido:")
    for item in desconocidos[:30]:
        print("-", item)

print("")
print("Proceso completado.")