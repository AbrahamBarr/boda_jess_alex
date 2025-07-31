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
invitaciones = df[df["Invitación dirigida a"].notna()]
boletos_dict = invitaciones.groupby("Invitación dirigida a")["Max Boletos"].max().to_dict()

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("invitacion.html", {
        "request": request,
        "event_date": "2025-11-15",
        "invitados": boletos_dict
    })

@app.post("/confirmar", response_class=HTMLResponse)
async def confirmar(request: Request, nombre: str = Form(...), asistentes: int = Form(...)):
    # Guardar en archivo CSV
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