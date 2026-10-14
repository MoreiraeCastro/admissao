import asyncio
import json
import os
import re
import shutil
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import List
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from docx import Document
from num2words import num2words

import dominio
from models import CAMPOS_ADMISSAO, DadosAdmissao, GerarArquivosResponse

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
GENERATED_DIR = BASE_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = BASE_DIR / "template_contrato.docx"

BRASILIA_TZ = ZoneInfo("America/Sao_Paulo")

# Estado em memória dos jobs de geração (sem banco de dados, conforme escopo do projeto).
# Guarda apenas o nome-base usado nos arquivos baixáveis. Os arquivos em si ficam no disco
# até a limpeza diária (ver _limpeza_diaria_loop) apagar tudo de uma vez, de madrugada.
JOBS_META: dict[str, dict] = {}


def _slugify_nome(nome: str) -> str:
    sem_acentos = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9]+", "", sem_acentos).lower()
    return slug or "colaborador"


def _limpar_todos_os_jobs() -> None:
    for job_dir in GENERATED_DIR.iterdir():
        if job_dir.is_dir():
            shutil.rmtree(job_dir, ignore_errors=True)
    JOBS_META.clear()


async def _limpeza_diaria_loop() -> None:
    """Apaga todos os arquivos gerados uma vez por dia, às 3h no horário de Brasília.

    Não apagamos na hora do download: os arquivos ficam disponíveis no servidor
    durante o dia (para o RH baixar/reenviar quando precisar) e são varridos de
    uma vez só de madrugada, quando ninguém deve estar usando o sistema.
    """
    while True:
        agora = datetime.now(BRASILIA_TZ)
        proxima_execucao = agora.replace(hour=3, minute=0, second=0, microsecond=0)
        if proxima_execucao <= agora:
            proxima_execucao += timedelta(days=1)
        await asyncio.sleep((proxima_execucao - agora).total_seconds())
        _limpar_todos_os_jobs()


ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")

app = FastAPI(title="Admissão RH — API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def _iniciar_limpeza_diaria() -> None:
    asyncio.create_task(_limpeza_diaria_loop())

# --------------------------------------------------------------------------
# Extração de dados via IA (Anthropic)
# --------------------------------------------------------------------------

SYSTEM_PROMPT = """Você é um assistente especialista em extrair dados cadastrais de documentos brasileiros a partir de imagens ou PDFs, para fins de admissão de funcionários.

Você pode receber vários documentos na mesma solicitação, sem limite de quantidade — por exemplo RG, CPF, CNH, CIN, comprovante de residência (conta de luz, água, telefone ou internet), certidões e outros. Combine as informações de TODOS os documentos enviados num único resultado: dados pessoais normalmente vêm do documento de identificação, e o endereço normalmente vem do comprovante de residência (mas também pode vir do próprio documento de identificação, quando presente). Se dois documentos trouxerem valores diferentes para o mesmo campo, prefira o documento oficial de identificação para dados pessoais e o comprovante mais recente para o endereço.

Analise cuidadosamente os documentos fornecidos e devolva a resposta APENAS como um objeto JSON válido — sem markdown, sem texto antes ou depois, sem comentários — contendo exatamente estas chaves:

{
  "nome_completo": "",
  "cpf": "",
  "rg": "",
  "rg_orgao_emissor": "",
  "rg_uf_emissor": "",
  "rg_data_emissao": "",
  "data_nascimento": "",
  "sexo": "",
  "naturalidade_cidade": "",
  "naturalidade_uf": "",
  "nacionalidade": "",
  "pais_nascimento": "",
  "nome_mae": "",
  "nome_pai": "",
  "endereco_logradouro": "",
  "endereco_numero": "",
  "endereco_complemento": "",
  "endereco_bairro": "",
  "endereco_cidade": "",
  "endereco_uf": "",
  "endereco_cep": "",
  "estado_civil": "",
  "grau_instrucao": "",
  "numero_ctps": "",
  "serie_ctps": "",
  "banco_identificado": "",
  "conta_corrente": "",
  "digito_conta": "",
  "tipo_conta": "",
  "deficiente_fisico": "",
  "email": "",
  "ddd_telefone": "",
  "telefone": "",
  "ddd_contato_2": "",
  "contato_2": ""
}

Regras:
- Nunca invente ou deduza informação que não esteja visível nos documentos. Se um campo não for encontrado em nenhum documento, use string vazia "".
- Formate o CPF como 000.000.000-00 quando encontrado.
- Formate datas como DD/MM/AAAA.
- "sexo" deve ser "Masculino" ou "Feminino" quando identificável.
- "rg_orgao_emissor" é a sigla do órgão emissor (ex: "SSP", "DETRAN"), separada da UF, que vai em "rg_uf_emissor" (ex: "RJ").
- "naturalidade_cidade" e "naturalidade_uf" vêm separados (ex: cidade "Rio de Janeiro", uf "RJ"). UFs sempre como sigla de 2 letras.
- "nacionalidade" em texto para o contrato (ex: "brasileira"); "pais_nascimento" é o nome do país de nascimento (ex: "BRASIL"), só se o documento mostrar ou se a naturalidade for uma cidade brasileira.
- Os campos "endereco_*" vêm cada um separado (não junte tudo num só campo). "endereco_logradouro" inclui o tipo (ex: "Rua das Flores", "Avenida Brasil"). "endereco_cep" apenas números.
- "estado_civil": só preencha com base em certidão de casamento, averbação de divórcio/separação, certidão de óbito do cônjuge ou declaração de união estável — use o código: S=Solteiro, C=Casado, V=Viúvo, D=Divorciado, J=Separado judicialmente, U=União estável. Certidão de nascimento sozinha NÃO prova estado civil: deixe vazio.
- "grau_instrucao": só com diploma, certificado de conclusão ou histórico escolar — use o código: 1=Analfabeto, 2=Primário (fundamental I) incompleto, 3=Primário completo, 4=Ginásio (fundamental II) incompleto, 5=Ginásio completo, 6=2º grau (médio) incompleto, 7=2º grau completo, 8=Superior incompleto, 9=Superior completo, 13=Pós-graduação, 10=Mestrado, 11=Doutorado.
- "numero_ctps" e "serie_ctps" da Carteira de Trabalho, quando enviada.
- Dados bancários só de cartão, extrato ou comprovante bancário em nome do colaborador: "banco_identificado" com nome e número do banco (ex: "Itaú (341)"), "conta_corrente" só os dígitos da conta sem o dígito verificador, "digito_conta" o dígito verificador, "tipo_conta" com o código 2=Conta corrente, 3=Conta salário, 4=Conta poupança.
- "deficiente_fisico": "1" apenas se houver laudo médico de deficiência; caso contrário deixe vazio.
- Telefones: DDD (2 dígitos) separado do número, apenas dígitos.
- Responda SOMENTE com o objeto JSON."""

IMAGE_MEDIA_TYPES = {"image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif"}


# A API da Anthropic aceita até 32 MB por requisição (o base64 aumenta ~33%) e o
# nginx corta envios acima de 20 MB — o frontend já reduz as fotos antes de mandar.
LIMITE_TOTAL_BYTES = 19 * 1024 * 1024


def _build_content_block(content_type: str, data_b64: str) -> dict:
    if content_type == "application/pdf":
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": data_b64},
        }
    media_type = "image/jpeg" if content_type == "image/jpg" else content_type
    return {
        "type": "image",
        "source": {"type": "base64", "media_type": media_type, "data": data_b64},
    }


def _extract_json(text: str) -> dict:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return json.loads(match.group(0))
    raise ValueError("A IA não retornou um JSON válido.")


@app.post("/api/extrair")
async def extrair_dados(documentos: List[UploadFile] = File(...)):
    if not documentos:
        raise HTTPException(status_code=400, detail="Nenhum documento enviado.")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=503,
            detail=(
                "ANTHROPIC_API_KEY não configurada no servidor. Defina a variável "
                "de ambiente (ou arquivo .env) para habilitar a extração via IA."
            ),
        )

    import base64

    import anthropic

    content_blocks = []
    nao_suportados = []
    total_bytes = 0
    for upload in documentos:
        raw = await upload.read()
        if not raw:
            continue
        content_type = (upload.content_type or "").lower()
        if content_type != "application/pdf" and content_type not in IMAGE_MEDIA_TYPES:
            # Antes caía como "image/jpeg" e a IA recusava o pedido inteiro (ex.: HEIC do iPhone).
            nao_suportados.append(upload.filename or content_type or "arquivo")
            continue
        total_bytes += len(raw)
        b64 = base64.b64encode(raw).decode("utf-8")
        content_blocks.append({"type": "text", "text": f"Documento: {upload.filename or 'sem nome'}"})
        content_blocks.append(_build_content_block(content_type, b64))

    if nao_suportados:
        raise HTTPException(
            status_code=400,
            detail=(
                "Formato não suportado pela IA: " + ", ".join(nao_suportados)
                + ". Envie como JPG, PNG, WEBP ou PDF (no iPhone, tire print da foto ou exporte como JPG)."
            ),
        )
    if not content_blocks:
        raise HTTPException(status_code=400, detail="Os arquivos enviados estão vazios.")
    if total_bytes > LIMITE_TOTAL_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Os documentos somam {total_bytes / 1024 / 1024:.1f} MB — o limite por envio é de ~19 MB. Divida em dois envios.",
        )

    content_blocks.append(
        {
            "type": "text",
            "text": (
                f"Extraia os dados cadastrais dos {sum(1 for b in content_blocks if b['type'] != 'text')} "
                "documento(s) acima, combinando as informações de todos eles."
            ),
        }
    )

    client = anthropic.Anthropic(api_key=api_key)

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=2048,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content_blocks}],
        )
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Erro ao chamar a API da Anthropic: {exc}")

    texto_resposta = "".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    )

    try:
        dados_brutos = _extract_json(texto_resposta)
    except (ValueError, json.JSONDecodeError):
        preview = texto_resposta.strip()[:500] or "(resposta vazia)"
        raise HTTPException(
            status_code=502,
            detail=(
                "A IA retornou uma resposta que não pôde ser interpretada como JSON "
                f"(stop_reason={response.stop_reason}). Resposta recebida: {preview}"
            ),
        )

    dados = DadosAdmissao(**{k: str(dados_brutos.get(k) or "") for k in CAMPOS_ADMISSAO})
    # Já devolve conferido contra a planilha do Domínio: códigos (país, estado civil…), calculados e pendências.
    resposta = _resultado_json(dominio.preparar(dados.model_dump()))
    resposta["banco_identificado"] = str(dados_brutos.get("banco_identificado") or "")
    return resposta


def _resultado_json(resultado: dominio.Resultado) -> dict:
    return {
        "dados": resultado.dados,
        "calculados": resultado.calculados,
        "erros": resultado.erros,
        "avisos": resultado.avisos,
        "dicas": resultado.dicas,
    }


@app.get("/api/campos")
async def campos():
    """Regras de cada campo (posição no layout, obrigatoriedade, opções do dicionário) — o frontend monta o formulário daqui."""
    return dominio.metadados_campos()


@app.post("/api/validar")
async def validar(dados: DadosAdmissao):
    """Confere os dados contra a planilha do Domínio sem gerar nada (o frontend chama enquanto o RH digita)."""
    return _resultado_json(dominio.preparar(dados.model_dump()))


# --------------------------------------------------------------------------
# Geração de arquivos (contrato .docx + layout de importação do Domínio)
# --------------------------------------------------------------------------


def _replace_placeholders_in_paragraph(paragraph, contexto: dict) -> None:
    if not paragraph.runs:
        return
    texto_completo = "".join(run.text for run in paragraph.runs)
    if "{{" not in texto_completo:
        return
    novo_texto = texto_completo
    for chave, valor in contexto.items():
        novo_texto = novo_texto.replace("{{" + chave + "}}", str(valor))
    if novo_texto == texto_completo:
        return
    paragraph.runs[0].text = novo_texto
    for run in paragraph.runs[1:]:
        run.text = ""


def _juntar(*partes: str, sep: str = " ") -> str:
    return sep.join(p for p in partes if p)


ESTADO_CIVIL_EXTENSO = {
    "S": "solteiro(a)",
    "C": "casado(a)",
    "V": "viúvo(a)",
    "D": "divorciado(a)",
    "O": "em regime de concubinato",
    "J": "separado(a) judicialmente",
    "U": "em união estável",
}

MESES_PT = [
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
    "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def _formatar_moeda_brl(valor_str: str) -> str:
    try:
        valor = Decimal(str(valor_str).replace(",", "."))
    except (InvalidOperation, ValueError):
        return ""
    texto = f"{valor:,.2f}"
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def _valor_por_extenso_reais(valor_str: str) -> str:
    try:
        valor = Decimal(str(valor_str).replace(",", "."))
    except (InvalidOperation, ValueError):
        return ""
    reais = int(valor)
    centavos = int(round((valor - reais) * 100))
    partes = [f"{num2words(reais, lang='pt_BR')} {'real' if reais == 1 else 'reais'}"]
    if centavos:
        partes.append(f"{num2words(centavos, lang='pt_BR')} {'centavo' if centavos == 1 else 'centavos'}")
    return " e ".join(partes)


def _montar_campos_combinados(dados: DadosAdmissao) -> dict:
    """Strings de exibição (para o contrato), combinando os campos estruturados."""
    endereco_linha1 = _juntar(
        dados.endereco_logradouro,
        f", {dados.endereco_numero}" if dados.endereco_numero else "",
        sep="",
    )
    endereco_completo = " - ".join(
        p
        for p in [
            endereco_linha1,
            dados.endereco_complemento,
            dados.endereco_bairro,
            _juntar(dados.endereco_cidade, dados.endereco_uf, sep="/"),
            f"CEP {dados.endereco_cep}" if dados.endereco_cep else "",
        ]
        if p
    )
    hoje = datetime.now(BRASILIA_TZ)
    return {
        "endereco_completo": endereco_completo,
        "naturalidade": _juntar(dados.naturalidade_cidade, dados.naturalidade_uf, sep="/"),
        "rg_orgao_uf": _juntar(dados.rg_orgao_emissor, dados.rg_uf_emissor, sep="/"),
        "estado_civil_extenso": ESTADO_CIVIL_EXTENSO.get(dados.estado_civil, ""),
        "salario_formatado": _formatar_moeda_brl(dados.salario),
        "salario_extenso": _valor_por_extenso_reais(dados.salario),
        "dia_assinatura": str(hoje.day),
        "mes_assinatura": MESES_PT[hoje.month],
        "ano_assinatura": str(hoje.year),
    }


def gerar_contrato_docx(dados: DadosAdmissao, destino: Path) -> None:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=(
                f"Template não encontrado em {TEMPLATE_PATH.name}. Execute "
                "scripts/create_template.py ou adicione o modelo real da empresa."
            ),
        )
    doc = Document(str(TEMPLATE_PATH))
    contexto = {**dados.model_dump(), **_montar_campos_combinados(dados)}

    for paragraph in doc.paragraphs:
        _replace_placeholders_in_paragraph(paragraph, contexto)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_placeholders_in_paragraph(paragraph, contexto)

    doc.save(str(destino))


def gerar_layout_dominio(dados: dict, destino: Path) -> None:
    """Gera o "empregados.txt" exatamente como a macro "Gerar" da planilha oficial do
    Domínio: as 78 colunas de uma linha separadas por TAB, sem cabeçalho, em ANSI
    (cp1252, o que o `Print #` do VBA grava no Windows) e quebra de linha CRLF.
    """
    with destino.open("w", encoding="cp1252", errors="replace", newline="\r\n") as f:
        f.write("\t".join(dominio.linha(dados)) + "\n")


def gerar_conferencia_dominio(dados: dict, destino: Path) -> None:
    """Versão legível (com rótulos) das mesmas 78 posições do arquivo oficial,
    só para o RH conferir visualmente antes de importar — não é o arquivo que
    vai pro Domínio (esse é o layout_dominio.txt, cru e sem rótulos).
    """
    valores = dominio.linha(dados)
    linhas = [
        f"{p.pos:02d} - {p.rotulo}: {valor}"
        for p, valor in zip(dominio.POSICOES, valores)
        if valor
    ]
    destino.write_text("\n".join(linhas) + "\n", encoding="utf-8")


@app.post("/api/gerar", response_model=GerarArquivosResponse)
async def gerar_arquivos(dados: DadosAdmissao):
    resultado = dominio.preparar(dados.model_dump())
    if resultado.erros:
        # O TXT só sai quando bate 100% com a planilha do Domínio — senão a importação falha lá dentro.
        return JSONResponse(
            status_code=422,
            content={
                "detail": f"{len(resultado.erros)} campo(s) precisam de correção antes de gerar o arquivo do Domínio.",
                **_resultado_json(resultado),
            },
        )
    dados_finais = DadosAdmissao(**resultado.dados)

    job_id = uuid.uuid4().hex
    job_dir = GENERATED_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    docx_path = job_dir / "contrato_trabalho.docx"
    dominio_path = job_dir / "layout_dominio.txt"
    conferencia_path = job_dir / "conferencia_dominio.txt"

    gerar_contrato_docx(dados_finais, docx_path)
    gerar_layout_dominio(resultado.dados, dominio_path)
    gerar_conferencia_dominio(resultado.dados, conferencia_path)

    # Nome-base usado nos arquivos baixados: {nomedocolaborador}_{ddmmyyyy} (data de
    # geração dos documentos — a mesma convenção que será usada na pasta do cloud).
    base_nome = f"{_slugify_nome(dados_finais.nome_completo)}_{date.today().strftime('%d%m%Y')}"

    JOBS_META[job_id] = {"base_nome": base_nome}

    # URLs relativas (sem "/" na frente) de propósito: assim funcionam tanto na
    # raiz do domínio (Render, IP direto) quanto atrás de um proxy reverso numa
    # sub-rota (ex: portalmoreiraecastro.com.br/admissao/).
    return GerarArquivosResponse(
        job_id=job_id,
        docx_url=f"api/download/{job_id}/docx",
        dominio_url=f"api/download/{job_id}/dominio",
        conferencia_url=f"api/download/{job_id}/conferencia",
    )


@app.get("/api/download/{job_id}/{tipo}")
async def download_arquivo(job_id: str, tipo: str):
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise HTTPException(status_code=400, detail="job_id inválido.")

    meta = JOBS_META.get(job_id)
    if not meta:
        raise HTTPException(status_code=404, detail="Arquivo não encontrado (job expirado ou já apagado na limpeza diária).")

    job_dir = GENERATED_DIR / job_id
    if tipo == "docx":
        caminho = job_dir / "contrato_trabalho.docx"
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        nome_arquivo = f"{meta['base_nome']}_contrato.docx"
    elif tipo == "dominio":
        caminho = job_dir / "layout_dominio.txt"
        media_type = "text/plain"
        nome_arquivo = f"{meta['base_nome']}_dominio.txt"
    elif tipo == "conferencia":
        caminho = job_dir / "conferencia_dominio.txt"
        media_type = "text/plain"
        nome_arquivo = f"{meta['base_nome']}_conferencia.txt"
    else:
        raise HTTPException(status_code=404, detail="Tipo de arquivo desconhecido.")

    if not caminho.exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado (job expirado ou inválido).")

    return FileResponse(caminho, media_type=media_type, filename=nome_arquivo)


# --------------------------------------------------------------------------
# Frontend estático
# --------------------------------------------------------------------------

app.mount("/", StaticFiles(directory=str(BASE_DIR / "static"), html=True), name="static")
