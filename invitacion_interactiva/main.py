from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
from datetime import datetime
import os
import json

os.makedirs("data", exist_ok=True)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Cargar datos de Excel (invitados)
df = pd.read_excel("data/invitados.xlsx")
df["Grupo"] = df["Invitación dirigida a"].ffill()
df["MaxBoletos_num"] = pd.to_numeric(df["Max Boletos"], errors="coerce").fillna(0).astype(int)

BOLETOS_POR_GRUPO = df.groupby("Grupo")["MaxBoletos_num"].max().to_dict()
NOMBRES_GRUPO = sorted([g for g in BOLETOS_POR_GRUPO.keys() if isinstance(g, str)])

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("invitacion.html", {
        "request": request,
        "nombres_grupo": NOMBRES_GRUPO,
        "limites_json": json.dumps(BOLETOS_POR_GRUPO, ensure_ascii=False),
        "event_date": "2025-11-15"
    })

@app.post("/confirmar", response_class=HTMLResponse)
async def confirmar(request: Request, nombre: str = Form(...), asistentes: int = Form(...)):
    max_permitido = int(BOLETOS_POR_GRUPO.get(nombre, 0))
    if asistentes > max_permitido:
        return templates.TemplateResponse("invitacion.html", {
            "request": request,
            "nombres_grupo": NOMBRES_GRUPO,
            "limites_json": json.dumps(BOLETOS_POR_GRUPO, ensure_ascii=False),
            "event_date": "2025-11-15",
            "error": f"El máximo permitido para {nombre} es {max_permitido}.",
            "nombre_valor": nombre,
            "asistentes_valor": max_permitido
        }, status_code=400)

    # === Guardar en Excel (data/confirmaciones.xlsx) ===
    excel_path = "data/confirmaciones.xlsx"
    fila = {
        "Nombre": nombre,
        "Asistentes": int(asistentes),
        "Fecha Confirmación": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    try:
        if os.path.exists(excel_path):
            df_conf = pd.read_excel(excel_path)
            df_conf = pd.concat([df_conf, pd.DataFrame([fila])], ignore_index=True)
        else:
            df_conf = pd.DataFrame([fila])

        df_conf.to_excel(excel_path, index=False)

    except Exception as e:
        # Fallback: si hubiera error al escribir Excel, guardamos en CSV
        import csv
        csv_path = "data/confirmaciones.csv"
        existe = os.path.isfile(csv_path)
        with open(csv_path, mode="a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not existe:
                writer.writerow(["Nombre", "Asistentes", "Fecha Confirmación"])
            writer.writerow([fila["Nombre"], fila["Asistentes"], fila["Fecha Confirmación"]])

    return templates.TemplateResponse("gracias.html", {
        "request": request,
        "nombre": nombre,
        "asistentes": asistentes
    })
