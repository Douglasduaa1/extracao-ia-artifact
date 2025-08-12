# =========================
# POC Extracao Open Source - DevTest Labs Artifact
# install.ps1
# =========================

$ErrorActionPreference = "Stop"
$projDir = "C:\POCExtract"

function Add-ToMachinePath([string]$path) {
  if (-not (Test-Path $path)) { return }
  $cur = [Environment]::GetEnvironmentVariable("Path","Machine")
  if ($null -eq $cur) { $cur = "" }
  if ($cur -notlike "*$path*") {
    [Environment]::SetEnvironmentVariable("Path", ($cur.TrimEnd(';') + ";" + $path), "Machine")
  }
}

function Invoke-Download($url, $out) {
  Write-Host "Baixando $url ..."
  Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
}

New-Item -ItemType Directory -Force -Path $projDir | Out-Null
Set-Location $projDir

# --- 1) Python 3.11 (All Users + PATH)
$pyUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
$pyExe = "$env:TEMP\python-3.11.9-amd64.exe"
Invoke-Download $pyUrl $pyExe
Start-Process -FilePath $pyExe -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_launcher=1" -Wait

# --- 2) Git
$gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.46.0.windows.1/Git-2.46.0-64-bit.exe"
$gitExe = "$env:TEMP\git-setup.exe"
Invoke-Download $gitUrl $gitExe
Start-Process -FilePath $gitExe -ArgumentList "/VERYSILENT /NORESTART" -Wait

# --- 3) Ghostscript
$gsUrl = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10031/gs10031w64.exe"
$gsExe = "$env:TEMP\gs-setup.exe"
Invoke-Download $gsUrl $gsExe
Start-Process -FilePath $gsExe -ArgumentList "/S" -Wait

# --- 4) QPDF
$qpdfUrl = "https://sourceforge.net/projects/qpdf/files/latest/download"
$qpdfZip = "$env:TEMP\qpdf.zip"
Invoke-Download $qpdfUrl $qpdfZip
Expand-Archive -Path $qpdfZip -DestinationPath "C:\Program Files\qpdf" -Force
$qpdfBin = (Get-ChildItem "C:\Program Files\qpdf" -Directory | Select-Object -First 1).FullName + "\bin"
Add-ToMachinePath $qpdfBin

# --- 5) Tesseract (UB Mannheim)
$tesUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.0.20221222/tesseract-ocr-w64-setup-5.3.0.20221222.exe"
$tesExe = "$env:TEMP\tesseract-setup.exe"
Invoke-Download $tesUrl $tesExe
Start-Process -FilePath $tesExe -ArgumentList "/VERYSILENT /NORESTART" -Wait
$tesPath = "C:\Program Files\Tesseract-OCR"
Add-ToMachinePath $tesPath

# --- Atualizar PATH no processo atual
$env:Path = [Environment]::GetEnvironmentVariable("Path","Machine")

# --- Verificações rápidas (não falha a instalação se der erro aqui)
try { python --version } catch {}
try { git --version } catch {}
try { tesseract --version } catch {}
try { qpdf --version } catch {}

# --- 6) Ambiente Python e deps
python -m venv "$projDir\.venv"
& "$projDir\.venv\Scripts\python.exe" -m pip install --upgrade pip

@'
fastapi==0.115.0
uvicorn[standard]==0.30.0
streamlit==1.37.0
pydantic==2.8.2
pdfplumber==0.11.4
ocrmypdf==16.0.5
Pillow==10.4.0
reportlab==3.6.13
'@ | Out-File -FilePath "$projDir\requirements.txt" -Encoding UTF8

& "$projDir\.venv\Scripts\python.exe" -m pip install -r "$projDir\requirements.txt"

# --- 7) API
@'
from fastapi import FastAPI, UploadFile, File
import tempfile, subprocess, uuid, re
import pdfplumber

app = FastAPI(title="OpenSource Extract POC (Windows)")

def run_ocrmypdf(src, dst):
    subprocess.check_call(["ocrmypdf", "--force-ocr", "--rotate-pages", "--deskew", "--output-type", "pdf", src, dst])

def extract_fields(text: str):
    cnpj = re.search(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", text)
    cpf  = re.search(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b", text)
    valor = re.search(r"(?i)(valor(?:\s+do\s+documento)?|total)[:\s]*R?\$?\s*([\d\.\,]+)", text)
    venc = re.search(r"(?i)(venc(?:imento)?)[:\s]*([0-3]?\d/[0-1]?\d/\d{2,4})", text)
    return {
        "cnpj": cnpj.group(0) if cnpj else None,
        "cpf": cpf.group(0) if cpf else None,
        "valor": valor.group(2) if valor else None,
        "vencimento": venc.group(2) if venc else None
    }

@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    data = await file.read()
    with tempfile.TemporaryDirectory() as td:
        raw = f"{td}/{uuid.uuid4()}.pdf"
        proc = f"{td}/proc.pdf"
        with open(raw, "wb") as f:
            f.write(data)

        run_ocrmypdf(raw, proc)

        pages_text = []
        with pdfplumber.open(proc) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                pages_text.append(txt)
        full_text = "\n".join(pages_text)
        fields = extract_fields(full_text)
        return {"ok": True, "fields": fields, "chars": len(full_text)}
'@ | Out-File -FilePath "$projDir\api.py" -Encoding UTF8

# --- 8) UI
@'
import streamlit as st
import requests, pandas as pd

st.set_page_config(page_title="POC Extração Open Source (Windows)", layout="wide")
st.title("🧾 POC — Extração de Campos (Open Source / Windows)")

api = st.text_input("URL da API", "http://localhost:8000/extract")
up = st.file_uploader("Envie um PDF", type=["pdf"])

if st.button("Extrair", type="primary") and up:
    with st.spinner("Processando..."):
        r = requests.post(api, files={"file": (up.name, up.getvalue(), "application/pdf")})
    if not r.ok:
        st.error(f"Erro: {r.status_code} {r.text}")
    else:
        j = r.json()
        st.success("Extração concluída")
        st.subheader("Campos")
        st.table(pd.DataFrame([j.get("fields", {})]))
        st.caption(f"Total de caracteres lidos: {j.get('chars')}")
'@ | Out-File -FilePath "$projDir\ui.py" -Encoding UTF8

# --- 9) Gerar PDFs de exemplo (via reportlab)
@'
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import date

samples_dir = r"C:\POCExtract\samples"
import os
os.makedirs(samples_dir, exist_ok=True)

def make_pdf(path, lines):
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    y = height - 72
    for line in lines:
        c.drawString(72, y, line)
        y -= 18
    c.showPage()
    c.save()

make_pdf(samples_dir + r"\boleto_bom.pdf", [
    "Empresa XYZ LTDA",
    "CNPJ: 12.345.678/0001-90",
    "CPF do pagador: 123.456.789-10",
    "Valor do Documento: R$ 149,43",
    "Vencimento: 25/08/2025"
])

make_pdf(samples_dir + r"\boleto_ruim.pdf", [
    "Empresa ABC S.A.",
    "CNPJ 98.765.432/0001-00",
    "TOTAL: R$ 2.345,67",
    "Venc.: 07/09/2025"
])

make_pdf(samples_dir + r"\fatura.pdf", [
    "Cliente: Fulano de Tal",
    "CPF: 987.654.321-00",
    "Valor Total: R$ 1.234,56",
    "Data de Vencimento: 15/09/2025"
])

print("Samples OK")
'@ | Out-File -FilePath "$projDir\make_samples.py" -Encoding UTF8

& "$projDir\.venv\Scripts\python.exe" "$projDir\make_samples.py"

# --- 10) Atalhos .bat
$desktop = [Environment]::GetFolderPath('Desktop')

@"
@echo off
cd /d $projDir
call .\.venv\Scripts\activate
uvicorn api:app --host 0.0.0.0 --port 8000
"@ | Out-File -FilePath "$desktop\Start_API.bat" -Encoding ASCII

@"
@echo off
cd /d $projDir
call .\.venv\Scripts\activate
streamlit run ui.py --server.port 8501 --server.address 0.0.0.0
"@ | Out-File -FilePath "$desktop\Start_UI.bat" -Encoding ASCII

# --- 11) Firewall
netsh advfirewall firewall add rule name="POC API 8000" dir=in action=allow protocol=TCP localport=8000 | Out-Null
netsh advfirewall firewall add rule name="POC UI 8501" dir=in action=allow protocol=TCP localport=8501 | Out-Null

Write-Host "Instalação concluída."
