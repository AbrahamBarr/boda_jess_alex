from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
from datetime import datetime
import csv
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Cargar datos de Excel
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

 # Guardar en CSV
    archivo = "data/confirmaciones.csv"
    existe = os.path.isfile(archivo)
    with open(archivo, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not existe:
            writer.writerow(["Nombre", "Asistentes", "Fecha Confirmación"])
        writer.writerow([nombre, asistentes, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

    return templates.TemplateResponse("gracias.html", {
        "request": request,
        "nombre": nombre
    })
