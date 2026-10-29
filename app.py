import csv
import io
import os
import re
import uuid
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI(
    title="Lotmax - Conversor de Veículos e Pneus",
    description="Recebe Excel, converte para CSV e disponibiliza uma URL para o Bubble.",
    version="2.2.0",
)

UPLOAD_DIR = Path(os.environ.get("LOTMAX_UPLOAD_DIR", "/tmp/lotmax-veiculos"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
MAX_FILE_SIZE = 20 * 1024 * 1024

INTEGER_FIELDS = {"11-Locacao_$"}
DECIMAL_FIELDS = {"11-Locacao_CFD"}

PNEUS_HEADERS = [
    "00-Marca",
    "00-Pneu_tipo",
    "00-Recapadora",
    "01-Codigo_aplicado",
    "01-Condicao",
    "01-Recapes",
    "01-Vida_util_atual",
    "01-Vida_util_fabricante",
    "01-Vida_util_recapado",
    "02-DOT_Fabricado",
    "02-Medida",
    "02-Valor_compra",
    "04-Local",
]


@app.get("/")
def health_check():
    return {
        "status": "ok",
        "service": "lotmax-conversor-veiculos",
        "version": "2.2.0",
        "message": "Serviço ativo. Envie um Excel para POST /veiculos ou POST /pneus."
    }


def normalizar_cabecalho(valor):
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar_numero(valor, campo):
    if valor is None or str(valor).strip() == "":
        return ""
    if isinstance(valor, bool):
        raise ValueError(f"Valor inválido no campo {campo}: {valor}")
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        numero = Decimal(texto)
    except InvalidOperation:
        raise ValueError(f"Valor numérico inválido no campo {campo}: {valor}")
    if not numero.is_finite():
        raise ValueError(f"Valor inválido no campo {campo}: {valor}")
    if campo in INTEGER_FIELDS:
        if numero != numero.to_integral_value():
            raise ValueError(
                f"O campo inteiro {campo} recebeu um valor fracionário: {valor}"
            )
        return str(int(numero))
    return format(numero, "f")


def ler_upload_xlsx(file):
    # Mantém a validação de extensão e tamanho usada nos dois endpoints.
    return None


@app.post("/veiculos")
async def receber_excel_veiculos(
    request: Request,
    file: UploadFile = File(...)
):
    original_name = file.filename or "veiculos.xlsx"
    if Path(original_name).suffix.lower() != ".xlsx":
        await file.close()
        raise HTTPException(status_code=400, detail="Envie uma planilha Excel no formato .xlsx.")
    conteudo = await file.read()
    await file.close()
    if not conteudo:
        raise HTTPException(status_code=400, detail="O arquivo está vazio.")
    if len(conteudo) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="O arquivo excede o limite de 20 MB.")

    workbook = None
    try:
        workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
        worksheet = workbook.active
        linhas = worksheet.iter_rows(values_only=True)
        cabecalhos_originais = next(linhas, None)
        if not cabecalhos_originais:
            raise ValueError("A planilha não contém cabeçalhos.")
        cabecalhos = [str(v).strip() if v is not None else "" for v in cabecalhos_originais]
        normalizados = [normalizar_cabecalho(c) for c in cabecalhos]
        if not any(cabecalhos):
            raise ValueError("A primeira linha está vazia.")
        indices_placa = [
            i for i, campo in enumerate(normalizados)
            if campo == "placa" or campo.endswith("-placa")
        ]
        if not indices_placa:
            raise ValueError("A coluna Placa não foi encontrada nos cabeçalhos.")
        indice_placa = indices_placa[0]

        saida = io.StringIO(newline="")
        escritor = csv.writer(saida, delimiter=",", quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
        escritor.writerow(cabecalhos)
        quantidade_linhas = 0
        linhas_descartadas = 0

        for numero_linha, linha in enumerate(linhas, start=2):
            valores = list(linha[:len(cabecalhos)])
            valores += [""] * (len(cabecalhos) - len(valores))
            placa = valores[indice_placa]
            if placa is None or str(placa).strip() == "":
                linhas_descartadas += 1
                continue
            linha_csv = []
            for campo, valor in zip(cabecalhos, valores):
                if campo in INTEGER_FIELDS or campo in DECIMAL_FIELDS:
                    try:
                        valor = normalizar_numero(valor, campo)
                    except ValueError as erro:
                        raise ValueError(f"Linha {numero_linha}: {erro}")
                elif valor is None:
                    valor = ""
                elif hasattr(valor, "isoformat"):
                    valor = valor.isoformat(sep=" ") if hasattr(valor, "hour") else valor.isoformat()
                else:
                    valor = str(valor)
                linha_csv.append(valor)
            escritor.writerow(linha_csv)
            quantidade_linhas += 1

        if quantidade_linhas == 0:
            raise ValueError("Nenhum veículo válido encontrado: todas as placas estão vazias.")
        csv_bytes = saida.getvalue().encode("utf-8-sig")
        stored_name = f"{uuid.uuid4().hex}.csv"
        (UPLOAD_DIR / stored_name).write_bytes(csv_bytes)
    except Exception as erro:
        if isinstance(erro, HTTPException):
            raise
        raise HTTPException(status_code=400, detail=f"Não foi possível converter o Excel: {erro}")
    finally:
        if workbook is not None:
            workbook.close()

    download_url = str(request.base_url).rstrip("/") + f"/arquivos/{stored_name}"
    return {
        "status": "ok",
        "filename": original_name,
        "csv_filename": "veiculos.csv",
        "size_bytes": len(csv_bytes),
        "rows": quantidade_linhas,
        "rows_discarded": linhas_descartadas,
        "url": download_url,
        "message": "Excel convertido em CSV. Linhas sem placa foram descartadas."
    }


@app.post("/pneus")
async def receber_excel_pneus(
    request: Request,
    file: UploadFile = File(...)
):
    original_name = file.filename or "pneus.xlsx"
    if Path(original_name).suffix.lower() != ".xlsx":
        await file.close()
        raise HTTPException(status_code=400, detail="Envie uma planilha Excel no formato .xlsx.")
    conteudo = await file.read()
    await file.close()
    if not conteudo:
        raise HTTPException(status_code=400, detail="O arquivo está vazio.")
    if len(conteudo) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="O arquivo excede o limite de 20 MB.")

    workbook = None
    try:
        workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
        if "Lista_Pneus" not in workbook.sheetnames:
            raise ValueError("A aba 'Lista_Pneus' não foi encontrada no arquivo.")
        worksheet = workbook["Lista_Pneus"]
        linhas = worksheet.iter_rows(values_only=True)
        cabecalhos_originais = next(linhas, None)
        if not cabecalhos_originais:
            raise ValueError("A aba 'Lista_Pneus' não contém cabeçalhos.")
        cabecalhos = [str(v).strip() if v is not None else "" for v in cabecalhos_originais]
        if cabecalhos != PNEUS_HEADERS:
            raise ValueError(
                "Os cabeçalhos da aba 'Lista_Pneus' não correspondem à estrutura esperada "
                "ou estão fora da ordem correta."
            )

        saida = io.StringIO(newline="")
        escritor = csv.writer(saida, delimiter=",", quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
        escritor.writerow(cabecalhos)
        quantidade_linhas = 0
        linhas_descartadas = 0

        for linha in linhas:
            valores = list(linha[:len(cabecalhos)])
            valores += [""] * (len(cabecalhos) - len(valores))
            # Pneus: ignorar apenas linhas completamente vazias; não exigir placa nem marca.
            if all(v is None or str(v).strip() == "" for v in valores):
                linhas_descartadas += 1
                continue
            linha_csv = []
            for valor in valores:
                if valor is None:
                    valor = ""
                elif hasattr(valor, "isoformat"):
                    valor = valor.isoformat(sep=" ") if hasattr(valor, "hour") else valor.isoformat()
                else:
                    valor = str(valor)
                linha_csv.append(valor)
            escritor.writerow(linha_csv)
            quantidade_linhas += 1

        if quantidade_linhas == 0:
            raise ValueError("Nenhum registro de pneu foi encontrado na aba 'Lista_Pneus'.")
        csv_bytes = saida.getvalue().encode("utf-8-sig")
        stored_name = f"{uuid.uuid4().hex}.csv"
        (UPLOAD_DIR / stored_name).write_bytes(csv_bytes)
    except Exception as erro:
        if isinstance(erro, HTTPException):
            raise
        raise HTTPException(status_code=400, detail=f"Não foi possível converter o Excel de pneus: {erro}")
    finally:
        if workbook is not None:
            workbook.close()

    download_url = str(request.base_url).rstrip("/") + f"/arquivos/{stored_name}?tipo=pneus"
    return {
        "status": "ok",
        "filename": original_name,
        "csv_filename": "pneus.csv",
        "sheet": "Lista_Pneus",
        "size_bytes": len(csv_bytes),
        "rows": quantidade_linhas,
        "rows_discarded": linhas_descartadas,
        "url": download_url,
        "message": "Excel convertido em CSV usando somente a aba Lista_Pneus."
    }


@app.get("/arquivos/{filename}")
def baixar_csv(filename: str, tipo: str = "veiculos"):
    if not re.fullmatch(r"[0-9a-f]{32}\.csv", filename):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")
    path = UPLOAD_DIR / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado ou expirado.")
    nome_download = "pneus.csv" if tipo == "pneus" else "veiculos.csv"
    return FileResponse(
        path,
        media_type="text/csv; charset=utf-8",
        filename=nome_download
    )
