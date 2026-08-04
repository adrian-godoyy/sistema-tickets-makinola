from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

ARCHIVO_CSV = BASE_DIR / "datos" / "tickets.csv"

print("Buscando archivo en:", ARCHIVO_CSV)
print("¿Existe el archivo?:", ARCHIVO_CSV.exists())

tickets = pd.read_csv(ARCHIVO_CSV)

print("====================================promedio de hora dia, semana y persona")
print("    IA DE ANÁLISIS DE TICKETS")
print("====================================")

print("\nTotal de tickets:", len(tickets))

print("\n--- TICKETS POR ESTADO ---")
print(tickets["estado"].value_counts())

print("\n--- TICKETS POR TÉCNICO ---")
print(tickets["tecnico"].value_counts())

print("\n--- TICKETS POR CATEGORÍA ---")
print(tickets["categoria"].value_counts())

promedio = tickets["tiempo_resolucion_horas"].mean()

print(
    "\nTiempo promedio de resolución:",
    round(promedio, 2),
    "horas"
)