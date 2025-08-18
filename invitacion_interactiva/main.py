from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
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

    saved_to = None
    try:
        if os.path.exists(excel_path):
            df_conf = pd.read_excel(excel_path)
            df_conf = pd.concat([df_conf, pd.DataFrame([fila])], ignore_index=True)
        else:
            df_conf = pd.DataFrame([fila])

        df_conf.to_excel(excel_path, index=False)
        saved_to = "xlsx"

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
        saved_to = "csv"
        # Log del error del Excel (visible en Render Logs)
        print(f"[CONFIRMACION][ERROR_XLSX] {e}", flush=True)

    # Log de confirmación (visible en Render Logs)
    print(f"[CONFIRMACION] nombre={nombre} asistentes={asistentes} saved_to={saved_to} at={datetime.now()}", flush=True)

    return templates.TemplateResponse("gracias.html", {
        "request": request,
        "nombre": nombre,
        "asistentes": asistentes
    })

@app.get("/admin/confirmaciones", response_class=HTMLResponse)
def admin_confirmaciones(request: Request):
    excel = "data/confirmaciones.xlsx"
    csvf = "data/confirmaciones.csv"

    # Intenta Excel primero; si no, CSV; si nada, DataFrame vacío
    df = None
    if os.path.exists(excel):
        try:
            df = pd.read_excel(excel)
        except Exception:
            df = None
    if df is None:
        if os.path.exists(csvf):
            import csv
            with open(csvf, "r", encoding="utf-8") as f:
                first = f.readline()
            header = 0 if first.strip().startswith("Nombre,") else None
            df = pd.read_csv(csvf, header=header)
            if header is None:
                df.columns = ["Nombre", "Asistentes", "Fecha Confirmación"]
        else:
            df = pd.DataFrame(columns=["Nombre", "Asistentes", "Fecha Confirmación"])

    if "Asistentes" in df.columns:
        df["Asistentes"] = pd.to_numeric(df["Asistentes"], errors="coerce").fillna(0).astype(int)

    total_registros = len(df)
    total_asistentes = int(df["Asistentes"].sum()) if "Asistentes" in df.columns else 0

    filas = []
    for _, r in df.fillna("").iterrows():
        filas.append(f"<tr><td>{r.get('Nombre','')}</td><td>{r.get('Asistentes','')}</td><td>{r.get('Fecha Confirmación','')}</td></tr>")

    html = f"""
    <html><head><meta charset="utf-8"><title>Confirmaciones</title>
    <style>
      body{{font-family:Arial,sans-serif; padding:20px;}}
      table{{border-collapse:collapse; width:100%;}}
      th,td{{border:1px solid #ddd; padding:8px; text-align:left;}}
      th{{background:#f4f4f4;}}
    </style>
    </head><body>
      <h2>Confirmaciones</h2>
      <p>Total registros: <b>{total_registros}</b> · Total asistentes: <b>{total_asistentes}</b></p>
      <p>
        <a href="/admin/descargar?formato=excel">Descargar Excel</a> ·
        <a href="/admin/descargar?formato=csv">Descargar CSV</a> ·
        <a href="/">Regresar</a>
      </p>
      <table>
        <thead><tr><th>Nombre</th><th>Asistentes</th><th>Fecha</th></tr></thead>
        <tbody>
          {''.join(filas) if filas else '<tr><td colspan="3">Sin registros aún.</td></tr>'}
        </tbody>
      </table>
    </body></html>
    """
    return HTMLResponse(html)

@app.get("/admin/descargar")
def descargar_confirmaciones(formato: str = "excel"):
    excel = "data/confirmaciones.xlsx"
    csvf = "data/confirmaciones.csv"

    if formato == "excel" and os.path.exists(excel):
        return FileResponse(excel, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="confirmaciones.xlsx")
    elif formato == "csv" and os.path.exists(csvf):
        return FileResponse(csvf, media_type="text/csv", filename="confirmaciones.csv")
    else:
        # Si no existe el solicitado, intenta el otro
        if os.path.exists(excel):
            return FileResponse(excel, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="confirmaciones.xlsx")
        if os.path.exists(csvf):
            return FileResponse(csvf, media_type="text/csv", filename="confirmaciones.csv")
        return HTMLResponse("<h3>No hay archivos de confirmaciones aún.</h3>", status_code=404)
