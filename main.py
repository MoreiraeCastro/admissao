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

SYSTEM_PROMPT = """Você extrai dados cadastrais de documentos brasileiros (fotos, PDFs e arquivos Word) para a admissão de um funcionário. O RH confere tudo depois, então é melhor deixar um campo vazio do que chutar.

Você pode receber vários documentos de uma vez. Combine as informações de TODOS num único resultado. Se dois documentos trouxerem valores diferentes para o mesmo campo, prefira o documento oficial de identificação para dados pessoais e o comprovante mais recente para o endereço, e registre a divergência em "duvidas".

REGRAS GERAIS
- Nunca invente nem deduza o que não está visível. Campo não encontrado = string vazia "".
- Copie números de documento EXATAMENTE como impressos, com os ZEROS À ESQUERDA (ex.: CTPS "0012345" fica "0012345", nunca "12345"). Vale para RG, CTPS, série, CNH, número do endereço, telefone e qualquer outro número.
- Datas sempre DD/MM/AAAA. CPF como 000.000.000-00. CEP só os 8 dígitos. UFs sempre a sigla de 2 letras.
- Nomes de pessoas como estão no documento oficial (com acentos), sem abreviar.
- Em "fontes", diga de qual arquivo (o nome que aparece em "Documento: ...") tirou cada campo preenchido.
- Em "duvidas", liste os campos que você preencheu sem certeza (foto borrada, texto cortado, dígito ilegível, documentos que divergem), com o motivo em poucas palavras. Não liste campos vazios.

ONDE ESTÁ CADA DADO, POR DOCUMENTO
- CIN (Carteira de Identidade Nacional, modelo novo, a partir de 2022): o número do documento É o próprio CPF. Preencha "cpf" e deixe "rg" vazio (o sistema usa o CPF como identidade). Órgão expedidor e UF aparecem como "SSP/RJ", "DETRAN/RJ" etc.
- RG antigo: o número do RG fica no campo "REGISTRO GERAL"; o órgão e a UF costumam vir juntos ("SSP/SP", "IFP-RJ", "DETRAN RJ") — separe a sigla em "rg_orgao_emissor" (ex.: "SSP", "IFP", "DETRAN") e a UF em "rg_uf_emissor". "Data de expedição" vai em "rg_data_emissao". Não confunda com a data de nascimento. Naturalidade aparece como "Cidade-UF".
- CNH: o RG aparece no campo "DOC. IDENTIDADE / ÓRG. EMISSOR / UF" (ex.: "123456789 DETRAN RJ") — o número vai em "rg" e o órgão/UF nos campos próprios; isso NÃO é o número da CNH (registro). Filiação traz pai e mãe. A data de emissão da CNH não é a data de expedição do RG: só use se não houver RG/CIN e o campo "DOC. IDENTIDADE" não trouxer data (nesse caso deixe "rg_data_emissao" vazio).
- CTPS física: "numero_ctps" é o número e "serie_ctps" a série (com a UF ao lado, que não entra na série). Mantenha os zeros à esquerda.
- CTPS Digital (tela ou PDF do app): não tem série — o número da CTPS digital são os 7 primeiros dígitos do CPF e a série são os 4 últimos dígitos do CPF; só preencha assim se o documento for claramente a CTPS Digital.
- Certidão de nascimento: nome, data de nascimento, naturalidade, filiação. Não prova estado civil.
- Certidão de casamento: estado civil "C" (ou "D"/"V"/"J" se houver averbação de divórcio, óbito ou separação). Se a pessoa mudou o nome no casamento, o "nome_completo" é o nome ATUAL (o que consta no documento de identidade mais recente ou o "passou a assinar").
- Comprovante de residência (luz, água, gás, telefone, internet, banco): use só o endereço. O titular da conta pode ser outra pessoa (pai, cônjuge, locador) — não use o nome dele para nada. Separe logradouro (com o tipo: "Rua", "Avenida"…), número, complemento (apto, bloco, casa, lote), bairro, cidade, UF e CEP.
- Título de eleitor, certificado de reservista, PIS/NIS: não preenchem nenhum campo daqui, exceto nome/data de nascimento/filiação para confirmar os outros documentos.
- Diploma, certificado ou histórico escolar: "grau_instrucao".
- Ficha cadastral ou formulário preenchido pelo candidato: pode trazer telefone, e-mail, estado civil e escolaridade — use, mas registre em "duvidas" o que só aparece ali.

CAMPOS ESPECÍFICOS
- "sexo": "Masculino" ou "Feminino", quando estiver no documento.
- "naturalidade_cidade" e "naturalidade_uf" separados (ex.: "Rio de Janeiro" e "RJ").
- "nacionalidade" em texto para o contrato (ex.: "brasileira"). Brasileiro nato: "brasileira".
- "estado_civil": só com certidão de casamento, averbação, certidão de óbito do cônjuge ou declaração de união estável — código: S=Solteiro, C=Casado, V=Viúvo, D=Divorciado, J=Separado judicialmente, U=União estável. Certidão de nascimento sozinha NÃO prova estado civil.
- "grau_instrucao": só com diploma, certificado de conclusão ou histórico — código: 1=Analfabeto, 2=Fundamental I incompleto, 3=Fundamental I completo, 4=Fundamental II incompleto, 5=Fundamental II completo, 6=Médio incompleto, 7=Médio completo, 8=Superior incompleto, 9=Superior completo, 13=Pós-graduação, 10=Mestrado, 11=Doutorado.
- "deficiente_fisico": "1" apenas se houver laudo médico de deficiência; senão vazio.
- Telefones: DDD (2 dígitos) separado do número; só dígitos."""

# Campos que a IA devolve (o resto é do RH ou calculado em dominio.py).
CAMPOS_IA = [c for c in CAMPOS_ADMISSAO if c not in ("pais_nascimento", "conta_corrente", "digito_conta", "tipo_conta")]

# Resposta com formato fixo (structured outputs): todos os campos sempre presentes e como texto
# (texto preserva zeros à esquerda), + de onde veio cada um e o que ficou em dúvida.
SCHEMA_EXTRACAO = {
    "type": "object",
    "properties": {
        "dados": {
            "type": "object",
            "properties": {campo: {"type": "string"} for campo in CAMPOS_IA},
            "required": CAMPOS_IA,
            "additionalProperties": False,
        },
        "fontes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"campo": {"type": "string", "enum": CAMPOS_IA}, "documento": {"type": "string"}},
                "required": ["campo", "documento"],
                "additionalProperties": False,
            },
        },
        "duvidas": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"campo": {"type": "string", "enum": CAMPOS_IA}, "motivo": {"type": "string"}},
                "required": ["campo", "motivo"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["dados", "fontes", "duvidas"],
    "additionalProperties": False,
}

# Erros do validador (dominio.preparar) que valem uma segunda leitura, e em qual campo lido pela IA eles caem.
CAMPO_LIDO_DO_ERRO = {"codigo_municipio_endereco": "endereco_cidade", "codigo_municipio_nascimento": "naturalidade_cidade"}


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


DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
# Imagens dentro do .docx (ex.: RG escaneado colado no Word) vão para a IA como imagem; a API aceita até ~5 MB cada.
IMAGEM_POR_EXTENSAO = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp"}
LIMITE_IMAGEM_DOCX = 5 * 1024 * 1024
LIMITE_DESCOMPACTADO_DOCX = 60 * 1024 * 1024  # proteção contra .docx "bomba" (zip que explode ao abrir)


def _eh_docx(content_type: str, nome: str) -> bool:
    return content_type == DOCX_MEDIA_TYPE or nome.lower().endswith(".docx")


def _blocos_docx(nome: str, raw: bytes) -> list[dict]:
    """Texto do Word (parágrafos, tabelas, cabeçalho e rodapé) + as imagens coladas nele, como blocos para a IA."""
    import base64
    import io
    import zipfile

    from docx import Document

    try:
        arquivo_zip = zipfile.ZipFile(io.BytesIO(raw))
        if sum(item.file_size for item in arquivo_zip.infolist()) > LIMITE_DESCOMPACTADO_DOCX:
            raise ValueError("grande demais")
        documento = Document(io.BytesIO(raw))
    except Exception as erro:  # zip inválido, .doc antigo renomeado, arquivo corrompido…
        raise HTTPException(status_code=400, detail=f"Não foi possível abrir “{nome}” como Word (.docx). Se for .doc antigo, salve como .docx ou PDF.") from erro

    linhas: list[str] = []
    for secao in documento.sections:
        linhas += [p.text for p in secao.header.paragraphs]
    linhas += [p.text for p in documento.paragraphs]
    for tabela in documento.tables:
        for linha in tabela.rows:
            celulas: list[str] = []
            for celula in linha.cells:
                texto = celula.text.strip()
                if texto and (not celulas or celulas[-1] != texto):  # células mescladas repetem o texto
                    celulas.append(texto)
            linhas.append(" | ".join(celulas))
    for secao in documento.sections:
        linhas += [p.text for p in secao.footer.paragraphs]
    texto = "\n".join(linha.strip() for linha in linhas if linha and linha.strip())

    blocos: list[dict] = [{"type": "text", "text": f"Documento: {nome} (Word)\n\n{texto or '(sem texto — veja as imagens abaixo)'}"}]
    ignoradas = 0
    for item in arquivo_zip.infolist():
        if not item.filename.startswith("word/media/"):
            continue
        media_type = IMAGEM_POR_EXTENSAO.get(os.path.splitext(item.filename)[1].lower())
        if not media_type or item.file_size > LIMITE_IMAGEM_DOCX:
            ignoradas += 1  # EMF/WMF/TIFF do Word ou imagem grande demais para a IA
            continue
        blocos.append({"type": "text", "text": f"Imagem dentro de {nome}"})
        blocos.append(_build_content_block(media_type, base64.b64encode(arquivo_zip.read(item)).decode("utf-8")))
    if ignoradas:
        blocos.append({"type": "text", "text": f"({ignoradas} imagem(ns) de {nome} não puderam ser lidas.)"})
    return blocos



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
        if _eh_docx(content_type, upload.filename or ""):
            total_bytes += len(raw)
            content_blocks += _blocos_docx(upload.filename or "documento.docx", raw)
            continue
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
                + ". Envie como JPG, PNG, WEBP, PDF ou Word .docx (no iPhone, tire print da foto ou exporte como JPG)."
            ),
        )
    if not content_blocks:
        raise HTTPException(status_code=400, detail="Os arquivos enviados estão vazios.")
    if total_bytes > LIMITE_TOTAL_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Os documentos somam {total_bytes / 1024 / 1024:.1f} MB — o limite por envio é de ~19 MB. Divida em dois envios.",
        )

    quantidade = sum(1 for b in content_blocks if b["type"] != "text") or 1
    # Cache no último documento: a segunda leitura reaproveita a mesma entrada (prompt + documentos) mais barato.
    content_blocks[-1] = {**content_blocks[-1], "cache_control": {"type": "ephemeral"}}
    client = anthropic.Anthropic(api_key=api_key)

    primeira = _ler_documentos(client, content_blocks, f"Extraia os dados cadastrais dos {quantidade} documento(s) acima, combinando as informações de todos eles.")
    dados_ia = {campo: primeira["dados"].get(campo, "") for campo in CAMPOS_IA}
    resultado = dominio.preparar(DadosAdmissao(**dados_ia).model_dump())

    # Segunda leitura: só os campos lidos que o validador recusou (CPF que não fecha, data impossível, cidade que não existe…).
    revisar = _campos_para_reler(resultado.erros, dados_ia)
    relidos: list[str] = []
    if revisar:
        lista = "\n".join(f"- {campo}: você leu “{dados_ia.get(campo, '')}”, mas a conferência apontou: {motivo}" for campo, motivo in revisar.items())
        try:
            segunda = _ler_documentos(client, content_blocks, (
                "Numa primeira leitura destes documentos, a conferência automática recusou os campos abaixo:\n"
                f"{lista}\n\nReleia os documentos com atenção SÓ para esses campos (dígito por dígito, zeros à esquerda incluídos). "
                "Devolva o resultado completo no mesmo formato; nos outros campos repita o que estiver nos documentos. "
                "Se o documento realmente mostra o valor recusado, mantenha-o e explique em \"duvidas\"."
            ))
            for campo in revisar:
                novo = segunda["dados"].get(campo, "")
                if novo != dados_ia.get(campo, ""):
                    dados_ia[campo] = novo
                    relidos.append(campo)
            primeira["duvidas"] = [d for d in primeira["duvidas"] if d["campo"] not in revisar] + [d for d in segunda["duvidas"] if d["campo"] in revisar]
            resultado = dominio.preparar(DadosAdmissao(**dados_ia).model_dump())
        except HTTPException:
            pass  # a segunda leitura é um bônus: se falhar, fica a primeira, com os erros apontados para o RH

    resposta = _resultado_json(resultado)
    resposta["ia"] = {
        "fontes": {f["campo"]: f["documento"] for f in primeira["fontes"] if dados_ia.get(f["campo"])},
        "duvidas": {d["campo"]: d["motivo"] for d in primeira["duvidas"]},
        "relidos": relidos,
    }
    return resposta


def _campos_para_reler(erros: dict[str, str], dados_ia: dict[str, str]) -> dict[str, str]:
    """Erros de conteúdo (não de campo em branco) em campos que a própria IA leu dos documentos."""
    revisar: dict[str, str] = {}
    for chave, erro in erros.items():
        campo = CAMPO_LIDO_DO_ERRO.get(chave, chave)
        if campo in CAMPOS_IA and dados_ia.get(campo) and not erro.startswith("Obrigatório"):
            revisar[campo] = erro
    return revisar


def _ler_documentos(client, content_blocks: list[dict], pedido: str) -> dict:
    """Uma leitura dos documentos com a resposta no formato fixo (SCHEMA_EXTRACAO)."""
    import anthropic

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=16000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": [*content_blocks, {"type": "text", "text": pedido}]}],
            # SDK 0.40 não conhece estes parâmetros: vão direto no corpo/cabeçalho da requisição.
            extra_body={"output_config": {"format": {"type": "json_schema", "schema": SCHEMA_EXTRACAO}}, "fallbacks": "default"},
            extra_headers={"anthropic-beta": "server-side-fallback-2026-07-01"},
        )
    except anthropic.APIError as exc:
        raise HTTPException(status_code=502, detail=f"Erro ao chamar a API da Anthropic: {exc}")
    if response.stop_reason == "refusal":
        raise HTTPException(status_code=502, detail="A IA se recusou a ler estes documentos. Confira se são documentos pessoais do colaborador e tente de novo.")
    if response.stop_reason == "max_tokens":
        raise HTTPException(status_code=502, detail="A resposta da IA foi cortada. Tente enviar menos documentos por vez.")
    texto = "".join(block.text for block in response.content if getattr(block, "type", None) == "text")
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        raise HTTPException(status_code=502, detail=f"A IA retornou uma resposta fora do formato (stop_reason={response.stop_reason}).")


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
