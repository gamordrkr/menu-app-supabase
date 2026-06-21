"""
Script para poblar Supabase con el catálogo de platillos de un cliente.

Uso:
    python poblar_supabase.py

Antes de correr, edita las variables SUPABASE_URL y SUPABASE_KEY abajo
(o ponlas como variables de entorno), y ajusta CLIENTE_NOMBRE si vas a
cargar otro cliente distinto a Stripe.
"""

import json
import os
from supabase import create_client

# --- Configuración -------------------------------------------------------

SUPABASE_URL = os.environ.get(
    "SUPABASE_URL", "https://eagzggevzabfstfbjbss.supabase.co"
)
SUPABASE_KEY = os.environ.get(
    "SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVhZ3pnZ2V2emFiZnN0ZmJqYnNzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE5ODg1NjAsImV4cCI6MjA5NzU2NDU2MH0.SbEZNRPRRHKXMnm4MUJRt4jJjiqXlb9ioDw4XI68Ppw",
)

CLIENTE_NOMBRE = "Stripe"
ARCHIVO_CATALOGO = "dishes_with_family.json"  # generado al limpiar el Excel original

# --- Script ----------------------------------------------------------------


def main():
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

    with open(ARCHIVO_CATALOGO, encoding="utf-8") as f:
        platillos = json.load(f)

    print(f"Platillos a cargar: {len(platillos)}")

    # 1. Verificar si el cliente ya existe (evita duplicados si corres el script 2 veces)
    slug = CLIENTE_NOMBRE.lower().strip().replace(" ", "_")
    existing = (
        supabase.table("clientes").select("*").eq("slug", slug).execute()
    )

    if existing.data:
        cliente_id = existing.data[0]["id"]
        print(f"Cliente '{CLIENTE_NOMBRE}' ya existe (id={cliente_id}).")
        confirm = input(
            "¿Borrar sus platillos actuales y volver a cargar? (s/n): "
        )
        if confirm.lower() == "s":
            supabase.table("platillos").delete().eq(
                "cliente_id", cliente_id
            ).execute()
            print("Platillos anteriores borrados.")
        else:
            print("Cancelado.")
            return
    else:
        result = (
            supabase.table("clientes")
            .insert({"nombre": CLIENTE_NOMBRE, "slug": slug})
            .execute()
        )
        cliente_id = result.data[0]["id"]
        print(f"Cliente '{CLIENTE_NOMBRE}' creado (id={cliente_id}).")

    # 2. Insertar platillos en lotes de 200 (evita payloads gigantes)
    rows = []
    for d in platillos:
        rows.append(
            {
                "cliente_id": cliente_id,
                "categoria": d["categoria"],
                "platillo": d["platillo"],
                "tipo": d.get("tipo", "") or "",
                "ingredientes": d.get("ingredientes", "") or "",
                "alergeno": d.get("alergeno", "") or "",
                "kcal": str(d.get("kcal", "")) if d.get("kcal") is not None else "",
                "familia_id": d["familia_id"],
            }
        )

    BATCH = 200
    total_insertados = 0
    for i in range(0, len(rows), BATCH):
        batch = rows[i : i + BATCH]
        supabase.table("platillos").insert(batch).execute()
        total_insertados += len(batch)
        print(f"  Insertados {total_insertados}/{len(rows)}...")

    print(f"\nListo. {total_insertados} platillos cargados para '{CLIENTE_NOMBRE}'.")
    print("Refresca la app en localhost:5173 y selecciona el cliente en el dropdown.")


if __name__ == "__main__":
    main()
