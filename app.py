from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from pathlib import Path
import tempfile
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse


from evidencias_core import generar_evidencias_desde_excel

app = FastAPI(title="Generador de evidencias QA")

@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
      <head>
        <title>Generador de evidencias QA</title>
        <style>
          body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            max-width: 600px;
            margin: 40px auto;
            padding: 20px;
            background: #0f172a;
            color: #e5e7eb;
          }
          h1 {
            text-align: center;
            margin-bottom: 0.5rem;
          }
          .card {
            background: #020617;
            border-radius: 12px;
            padding: 20px 24px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.5);
          }
          label {
            display: block;
            margin-top: 16px;
            font-size: 0.95rem;
          }
          input[type="text"],
          input[type="file"] {
            margin-top: 6px;
            width: 100%;
            padding: 8px 10px;
            border-radius: 8px;
            border: 1px solid #1f2937;
            background: #020617;
            color: #e5e7eb;
          }
          .row {
            display: flex;
            gap: 12px;
            margin-top: 12px;
            font-size: 0.9rem;
          }
          .row > label {
            flex: 1;
            display: flex;
            align-items: center;
            gap: 6px;
            margin-top: 0;
          }
          button {
            margin-top: 20px;
            width: 100%;
            padding: 10px 14px;
            border-radius: 999px;
            border: none;
            background: #22c55e;
            color: #020617;
            font-weight: 600;
            cursor: pointer;
          }
          button:hover {
            background: #16a34a;
          }
          .hint {
            font-size: 0.8rem;
            color: #9ca3af;
            margin-top: 4px;
          }
          a {
            color: #38bdf8;
          }
        </style>
      </head>
      <body>
        <h1>Generador de evidencias QA</h1>
        <p style="text-align:center;font-size:0.9rem;color:#9ca3af;">
          Sube el Excel de casos de prueba y descarga el Word con las evidencias.
        </p>
        <div class="card">
          <form action="/generar" method="post" enctype="multipart/form-data">
            <label>
              Archivo Excel de casos de prueba
              <input type="file" name="file" required />
            </label>

            <label>
              Nombre de la hoja
              <input type="text" name="hoja" value="Casos de Prueba" />
              <div class="hint">Debe coincidir con el nombre de la hoja en el Excel.</div>
            </label>

            <div class="row">
              <label>
                <input type="checkbox" name="sin_consolidado" />
                No generar consolidado
              </label>
              <label>
                <input type="checkbox" name="sin_individuales" checked />
                No generar individuales
              </label>
            </div>

            <button type="submit">Generar evidencias</button>
            <p class="hint">
              También puedes usar la API desde <a href="/docs" target="_blank">/docs</a>.
            </p>
          </form>
        </div>
      </body>
    </html>
    """


@app.post("/generar")
async def generar_evidencias(
    file: UploadFile = File(...),
    hoja: str = Form("Casos de Prueba"),
    sin_consolidado: bool = Form(False),
    sin_individuales: bool = Form(True),
):
    """
    Endpoint que recibe un Excel y devuelve el DOCX consolidado.
    """

    # ⚠️ Creamos un directorio temporal PERO SIN 'with'
    # para que no se borre antes de que FileResponse lo use.
    tmpdir = tempfile.mkdtemp()
    tmpdir_path = Path(tmpdir)

    # Guardar el Excel subido
    excel_path = tmpdir_path / file.filename
    contents = await file.read()
    excel_path.write_bytes(contents)

    # Carpeta de salida
    salida_dir = tmpdir_path / "evidencias_out"

    # Llamamos al motor
    generar_evidencias_desde_excel(
        excel_path=str(excel_path),
        hoja=hoja,
        salida=str(salida_dir),
        sin_consolidado=sin_consolidado,
        sin_individuales=sin_individuales,
        map_json=None,
    )

    consolidated = salida_dir / "Evidencias_Consolidadas.docx"

    if not consolidated.exists():
        return JSONResponse(
            status_code=500,
            content={"detail": f"No se generó el DOCX consolidado en {consolidated}."},
        )

    # Devolvemos el archivo como descarga
    return FileResponse(
        path=str(consolidated),
        filename=consolidated.name,
        media_type=(
            "application/"
            "vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )
