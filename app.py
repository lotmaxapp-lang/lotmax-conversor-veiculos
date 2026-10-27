import os
import re
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

app = FastAPI(
    title="Lotmax - Conversor de Veículos",
    description="Recebe um CSV já convertido pelo Lotmax e disponibiliza uma URL temporária para importação no Bubble.",
    version="1.0.0",
)

# No Render, este diretório é temporário: arquivos podem ser perdidos
# quando o serviço for reiniciado ou publicado novamente.
UPLOAD_DIR = Path(os.environ.get("LOTMAX_UPLOAD_DIR", "/tmp/lotmax-veiculos"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB


@app.get("/")
def health_check():
    return {
        "status": "ok",
        "service": "lotmax-conversor-veiculos",
        "message": "Serviço ativo. Envie um CSV para POST /veiculos.",
    }


@app.post("/veiculos")
async def receber_csv_veiculos(request: Request, file: UploadFile = File(...)):
    """Recebe um arquivo CSV e retorna uma URL para o Bubble baixá-lo."""
    original_name = file.filename or "veiculos.csv"

    if not original_name.lower().endswith(".csv"):
        raise HTTPException(status_code=400, detail="Envie um arquivo com extensão .csv.")

    stored_name = f"{uuid.uuid4().hex}.csv"
    destination = UPLOAD_DIR / stored_name
    total = 0

    try:
        with destination.open("wb") as output:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_FILE_SIZE:
                    destination.unlink(missing_ok=True)
                    raise HTTPException(
                        status_code=413,
                        detail="O CSV excede o limite de 20 MB.",
                    )
                output.write(chunk)
    finally:
        await file.close()

    if total == 0:
        destination.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail="O arquivo CSV está vazio.")

    download_url = str(request.base_url).rstrip("/") + f"/arquivos/{stored_name}"
    return {
        "status": "ok",
        "filename": original_name,
        "size_bytes": total,
        "url": download_url,
        "message": "CSV recebido. Use o campo 'url' na etapa de importação do Bubble.",
    }


@app.get("/arquivos/{filename}")
def baixar_csv(filename: str):
    """Disponibiliza o CSV recebido para download pelo Bubble."""
    if not re.fullmatch(r"[0-9a-f]{32}\.csv", filename):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")

    path = UPLOAD_DIR / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado ou expirado.")

    return FileResponse(
        path=path,
        media_type="text/csv; charset=utf-8",
        filename="veiculos.csv",
    )
