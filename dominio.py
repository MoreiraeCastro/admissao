"""Layout de importação de admissão do Domínio Web (empregados.txt) — fonte única das regras.

Tudo aqui foi conferido contra a planilha oficial do Domínio usada pelo escritório
(data/fonte/empregados_modelo_dominio.xls): abas "Preencher" (cabeçalho com * nos
obrigatórios), "Dicionário dos campos" (códigos válidos de cada posição), "Código
dos municípios" e "Código dos países", e a macro VBA "Gerar", que só junta as 78
colunas de cada linha com TAB (Print #, ANSI/cp1252, CRLF).

Fluxo: `preparar(dados)` normaliza o que o RH/IA digitaram (ponto decimal, só
dígitos, datas), calcula o que dá pra calcular (sem sobrescrever o que o RH
preencheu) e confere cada posição contra o dicionário → `linha(dados)` monta as
78 colunas na ordem exata da planilha.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent / "data"

with open(DATA_DIR / "municipios_dominio.json", encoding="utf-8") as _f:
    # {"RJ": {"NITEROI": 3638, ...}, ...} — por UF: 242 nomes de município se repetem entre estados.
    MUNICIPIOS: dict[str, dict[str, int]] = json.load(_f)
with open(DATA_DIR / "paises_dominio.json", encoding="utf-8") as _f:
    PAISES: dict[str, int] = json.load(_f)

UFS = sorted(MUNICIPIOS)
PAIS_BRASIL = "30"


def normalizar_nome(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", sem_acento).strip().upper()


def codigo_municipio(cidade: str, uf: str) -> str:
    """Código do Domínio para a cidade NAQUELA UF (busca exata, sem acento/caixa); "" se não achar."""
    codigo = MUNICIPIOS.get(normalizar_nome(uf), {}).get(normalizar_nome(cidade))
    return str(codigo) if codigo is not None else ""


# --------------------------------------------------------------------------
# Dicionário dos campos (opções válidas) — transcrito da aba "Dicionário dos campos"
# --------------------------------------------------------------------------

OPCOES: dict[str, list[tuple[str, str]]] = {
    "vinculo_empregaticio": [
        ("0", "Não aparece no SEFIP"), ("1", "Celetista"), ("2", "Estatutário"), ("3", "Trabalhador Avulso"),
        ("4", "Trabalhador Temporário Lei 6.019/74"), ("5", "Celetista Contrato Prazo Determinado"),
        ("6", "Trabalhador Rural"), ("7", "Diretor s/ Vínculo c/ FGTS"), ("8", "Servidor Público não Efetivo"),
        ("9", "Trab. Rural Contrato Prazo Determinado"), ("10", "Contrato Prazo Determinado Lei 9.601/98"),
        ("11", "Diretor sem vínculo, SEM FGTS"), ("13", "Autônomo"), ("14", "Aut. Classe 14"), ("15", "Transp. 15"),
        ("16", "Transp. Classe 16"), ("17", "Cooperado"), ("18", "Transportador cooperado"),
        ("22", "Contrib. indiv. contratado por contrib. indiv"),
        ("23", "Contrib. indiv. transportador aut. contratado por contrib. indiv."),
        ("24", "Cooperado que presta serv. para entidade beneficente"),
        ("25", "Transp. cooperado que presta serv. para entidade beneficente"),
        ("50", "Estagiário"), ("51", "Contrato Experiência"), ("52", "Empregado Doméstico"),
        ("53", "Menor Aprendiz"), ("54", "Empregado Doméstico Regime Tempo Parcial"),
        ("99", "Celetista regime tempo parcial"),
    ],
    "optante_fgts": [("N", "Não"), ("S", "Sim")],
    "ocorrencia_sefip": [
        ("0", "Nunca teve agente nocivo"), ("1", "01 - Não exposto a agente nocivo"),
        ("2", "02 - Exp. agente nocivo (Apo. Esp. 15 anos)"), ("3", "03 - Exp. agente nocivo (Apo. Esp. 20 anos)"),
        ("4", "04 - Exp. agente nocivo (Apo. Esp. 25 anos)"), ("5", "05 - Não exposto a agente nocivo"),
        ("6", "06 - Exp. agente nocivo (Apo. Esp. 15 anos)"), ("7", "07 - Exp. agente nocivo (Apo. Esp. 20 anos)"),
        ("8", "08 - Exp. agente nocivo (Apo. Esp. 25 anos)"),
    ],
    "tipo_conta": [("1", "Não informada"), ("2", "Conta corrente"), ("3", "Conta salário"), ("4", "Conta poupança")],
    "multiplos_vinculos": [("B", "Sim, abaixo do teto"), ("D", "Sim, neste banco"), ("N", "Não"), ("T", "Sim, acima do teto")],
    "pagou_contribuicao_sindical": [("L", "Prof. Liberal"), ("N", "Não"), ("S", "Sim"), ("X", "Não pagar")],
    "categoria": [
        ("1", "Mensalista"), ("3", "Semanalista"), ("4", "Diarista"), ("5", "Horista"),
        ("6", "Tarefeiro"), ("7", "Comissionado"), ("8", "Diretor"),
    ],
    "tipo_indicativo_admissao": [
        ("1", "1º emprego"), ("2", "Normal"), ("3", "Decorrente de ação fiscal"), ("4", "Decorrente de decisão judicial"),
    ],
    "grau_instrucao": [
        ("1", "Analfabeto"), ("2", "Primário incompleto"), ("3", "Primário completo"), ("4", "Ginásio incompleto"),
        ("5", "Ginásio completo"), ("6", "2º grau incompleto"), ("7", "2º grau completo"), ("8", "Superior incompleto"),
        ("9", "Superior completo"), ("10", "Mestrado"), ("11", "Doutorado"), ("12", "PhD"), ("13", "Pós-graduação"),
    ],
    "deficiente_fisico": [("0", "Não"), ("1", "Sim")],
    "tipo_admissao": [
        ("1", "Admissão"), ("2", "Admissão por sucessão, incorporação ou fusão"),
        ("3", "Transferência com ônus"), ("4", "Transferência sem ônus"), ("5", "Reintegração"),
    ],
    "forma_pagamento": [("C", "Crédito CC"), ("D", "Dinheiro"), ("H", "Cheque")],
    "situacao_qualificacao_cadastral": [("0", "Não qualificado"), ("1", "Qualificado")],
    "sexo": [("F", "Feminino"), ("M", "Masculino")],
    "raca_cor": [("0", "Indígena"), ("2", "Branca"), ("4", "Preta"), ("6", "Amarela"), ("8", "Parda"), ("9", "Não informada")],
    "estado_civil": [
        ("S", "Solteiro"), ("C", "Casado"), ("V", "Viúvo"), ("D", "Divorciado"), ("O", "Concubinato"),
        ("J", "Separado judicialmente"), ("U", "União estável"),
    ],
    "recebia_seguro_desemprego": [("0", "Não"), ("1", "Sim")],
    "tipo_endereco": [
        ("1", "Aeroporto"), ("2", "Alameda"), ("3", "Área"), ("4", "Avenida"), ("5", "Campo"), ("6", "Chácara"),
        ("7", "Colônia"), ("8", "Condomínio"), ("9", "Conjunto"), ("10", "Distrito"), ("11", "Esplanada"),
        ("12", "Estação"), ("13", "Estrada"), ("14", "Favela"), ("15", "Fazenda"), ("16", "Feira"), ("17", "Jardim"),
        ("18", "Ladeira"), ("19", "Lago"), ("20", "Lagoa"), ("21", "Largo"), ("22", "Loteamento"), ("23", "Morro"),
        ("24", "Núcleo"), ("25", "Parque"), ("26", "Passarela"), ("27", "Pátio"), ("28", "Praça"), ("29", "Quadra"),
        ("30", "Recanto"), ("31", "Residencial"), ("32", "Rodovia"), ("33", "Rua"), ("34", "Setor"), ("35", "Sítio"),
        ("36", "Travessa"), ("37", "Trecho"), ("38", "Trevo"), ("40", "Vereda"), ("41", "Via"), ("42", "Viaduto"),
        ("43", "Viela"), ("44", "Vila"), ("49", "Vale"), ("99", "Outro"),
    ],
    "contrato_experiencia": [("0", "Não possui"), ("1", "Possui")],
    # Não é coluna do layout — escolhe como o sistema calcula fim/prorrogação da experiência (o RH sempre pode digitar).
    "contagem_experiencia": [
        ("CLT", "Pela CLT, contando o 1º dia (30 dias a partir de 01/10 terminam em 30/10)"),
        ("MES", "Mês cheio, como o exemplo do Domínio (30 dias a partir de 01/10 terminam em 01/11)"),
    ],
    "clausula_assecuratoria": [
        ("1", "Regido pelo art. 479/480 CLT ou sem cláusula assecuratória"), ("2", "Regido pelo art. 481 CLT"),
    ],
    "forma_calculo_pagamento": [("1", "Por mês"), ("2", "Por semana")],
    "atividade_simples_nacional": [
        ("1", "Não optante"), ("2", "Optante enquadrada nos anexos I a III e V"),
        ("3", "Optante enquadrada no anexo IV"), ("4", "Optante com atividades concomitantes"),
    ],
}
OPCOES["pais_nacionalidade"] = OPCOES["pais_endereco"] = OPCOES["pais_nascimento"] = [
    (str(codigo), nome) for nome, codigo in sorted(PAISES.items(), key=lambda item: item[1])
]
for _chave in ("endereco_uf", "rg_uf_emissor", "naturalidade_uf"):
    OPCOES[_chave] = [(uf, uf) for uf in UFS]

# Categoria eSocial (posição 77) a partir do vínculo — tabela da própria posição 77 do dicionário.
# Só vínculos com UMA resposta inequívoca; o resto fica para o RH.
CATEGORIA_ESOCIAL_POR_VINCULO = {
    "1": "101", "3": "202", "4": "106", "5": "101", "6": "101", "10": "105",
    "52": "104", "53": "103", "54": "104", "99": "101",
}

# Tipo de endereço (posição 51) pela 1ª palavra do logradouro.
TIPO_ENDERECO_POR_PREFIXO = {
    "AEROPORTO": "1", "ALAMEDA": "2", "AL": "2", "AREA": "3", "AVENIDA": "4", "AV": "4", "CAMPO": "5",
    "CHACARA": "6", "COLONIA": "7", "CONDOMINIO": "8", "CONJUNTO": "9", "DISTRITO": "10", "ESPLANADA": "11",
    "ESTACAO": "12", "ESTRADA": "13", "ESTR": "13", "EST": "13", "FAVELA": "14", "FAZENDA": "15", "FEIRA": "16",
    "JARDIM": "17", "LADEIRA": "18", "LAGO": "19", "LAGOA": "20", "LARGO": "21", "LOTEAMENTO": "22", "MORRO": "23",
    "NUCLEO": "24", "PARQUE": "25", "PASSARELA": "26", "PATIO": "27", "PRACA": "28", "PCA": "28", "QUADRA": "29",
    "RECANTO": "30", "RESIDENCIAL": "31", "RODOVIA": "32", "ROD": "32", "RUA": "33", "R": "33", "SETOR": "34",
    "SITIO": "35", "TRAVESSA": "36", "TV": "36", "TRAV": "36", "TRECHO": "37", "TREVO": "38", "VEREDA": "40",
    "VIA": "41", "VIADUTO": "42", "VIELA": "43", "VILA": "44", "VALE": "49",
}

# Horas por semana → (horas dia, horas mês) nos formatos do próprio dicionário (7.33 | 8.00; 220.00 | 200.00).
HORAS_POR_SEMANA = {
    "44.00": ("7.33", "220.00"), "40.00": ("8.00", "200.00"), "36.00": ("6.00", "180.00"),
    "30.00": ("6.00", "150.00"), "25.00": ("5.00", "125.00"), "20.00": ("4.00", "100.00"),
}


# --------------------------------------------------------------------------
# As 78 posições — ordem, chave, rótulo, obrigatoriedade (* da aba Preencher) e tipo
# --------------------------------------------------------------------------

@dataclass(frozen=True)
class Posicao:
    pos: int
    chave: str  # campo de DadosAdmissao; "" = valor fixo/derivado (ver VALORES_FIXOS)
    rotulo: str
    obrigatorio: bool
    tipo: str  # texto | digitos | decimal | data | opcao | uf | email


POSICOES = [
    Posicao(1, "", "Código empresa", True, "digitos"),
    Posicao(2, "codigo_empregado", "Código empregado", True, "digitos"),
    Posicao(3, "nome_completo", "Nome", True, "texto"),
    Posicao(4, "cpf", "CPF", True, "digitos"),
    Posicao(5, "codigo_filial", "Código filial", True, "digitos"),
    Posicao(6, "codigo_servico", "Código serviço", True, "digitos"),
    Posicao(7, "codigo_cargo", "Código cargo", True, "digitos"),
    Posicao(8, "codigo_departamento", "Código departamento", True, "digitos"),
    Posicao(9, "codigo_centro_custos", "Código centro de custos", True, "digitos"),
    Posicao(10, "salario", "Salário", True, "decimal"),
    Posicao(11, "data_admissao", "Data da admissão", True, "data"),
    Posicao(12, "codigo_esocial_empregado", "Código do eSocial", False, "texto"),
    Posicao(13, "vinculo_empregaticio", "Vínculo empregatício", True, "opcao"),
    Posicao(14, "optante_fgts", "Optante do FGTS", True, "opcao"),
    Posicao(15, "data_opcao_fgts", "Data opção FGTS", False, "data"),
    Posicao(16, "ocorrencia_sefip", "Ocorrência SEFIP", True, "opcao"),
    Posicao(17, "endereco_logradouro", "Endereço", False, "texto"),
    Posicao(18, "endereco_numero", "Número endereço", False, "texto"),
    Posicao(19, "endereco_complemento", "Complemento endereço", False, "texto"),
    Posicao(20, "endereco_bairro", "Bairro", True, "texto"),
    Posicao(21, "endereco_cidade", "Cidade", True, "texto"),
    Posicao(22, "endereco_uf", "Sigla UF", True, "uf"),
    Posicao(23, "endereco_cep", "CEP", True, "digitos"),
    Posicao(24, "data_nascimento", "Data nascimento", True, "data"),
    # Banco (25 a 28) saiu do formulário (pedido do RH em 21/10/2026): conta "não informada" e banco em branco.
    Posicao(25, "codigo_banco", "Código banco", False, "digitos"),
    Posicao(26, "conta_corrente", "Número conta corrente", False, "digitos"),
    Posicao(27, "digito_conta", "Dígito conta corrente", False, "texto"),
    Posicao(28, "tipo_conta", "Tipo da conta", True, "opcao"),
    Posicao(29, "codigo_convenio_coletivo", "Código da convenção coletiva", True, "digitos"),
    Posicao(30, "multiplos_vinculos", "Múltiplos vínculos", True, "opcao"),
    Posicao(31, "pagou_contribuicao_sindical", "Pagou contr. sindical", True, "opcao"),
    Posicao(32, "vencimento_ferias", "Vencimento das férias", True, "data"),
    Posicao(33, "horas_dia", "Horas dia", True, "decimal"),
    Posicao(34, "categoria", "Categoria", True, "opcao"),
    Posicao(35, "tipo_indicativo_admissao", "Tipo de indicativo da admissão", True, "opcao"),
    Posicao(36, "grau_instrucao", "Grau instrução", True, "opcao"),
    Posicao(37, "deficiente_fisico", "Deficiente físico", True, "opcao"),
    Posicao(38, "tipo_admissao", "Tipo de admissão", True, "opcao"),
    Posicao(39, "horas_semana", "Horas semana", True, "decimal"),
    Posicao(40, "forma_pagamento", "Forma pagamento", True, "opcao"),
    Posicao(41, "horas_mes", "Horas mês", True, "decimal"),
    Posicao(42, "situacao_qualificacao_cadastral", "Situação da qualificação cadastral (INSS/RF)", True, "opcao"),
    Posicao(43, "sexo", "Sexo", True, "opcao"),
    Posicao(44, "pais_nacionalidade", "País nacionalidade", True, "opcao"),
    Posicao(45, "raca_cor", "Raça/Cor", True, "opcao"),
    Posicao(46, "estado_civil", "Estado civil", True, "opcao"),
    Posicao(47, "recebia_seguro_desemprego", "Recebia seguro desemprego", True, "opcao"),
    Posicao(48, "codigo_jornada", "Código da jornada", True, "digitos"),
    Posicao(49, "nome_pai", "Nome do pai", False, "texto"),
    Posicao(50, "nome_mae", "Nome da mãe", False, "texto"),
    Posicao(51, "tipo_endereco", "Tipo de endereço", True, "opcao"),
    Posicao(52, "codigo_municipio_endereco", "Código do município do endereço", True, "digitos"),
    Posicao(53, "codigo_municipio_nascimento", "Código do município de nascimento", True, "digitos"),
    Posicao(54, "pais_endereco", "Código do país do endereço", True, "opcao"),
    Posicao(55, "pais_nascimento", "Código do país de nascimento", True, "opcao"),
    Posicao(56, "rg", "Número da identidade (RG)", False, "texto"),
    Posicao(57, "rg_orgao_emissor", "Órgão de expedição do RG", False, "texto"),
    Posicao(58, "rg_uf_emissor", "UF de expedição do RG", False, "uf"),
    Posicao(59, "rg_data_emissao", "Data de expedição do RG", False, "data"),
    Posicao(60, "contrato_experiencia", "Contrato de experiência", False, "opcao"),
    Posicao(61, "clausula_assecuratoria", "Cláusula assecuratória", False, "opcao"),
    Posicao(62, "dias_contrato_experiencia", "Dias do contrato de experiência", False, "digitos"),
    Posicao(63, "inicio_prazo_determinado", "Início prazo determinado", False, "data"),
    Posicao(64, "fim_prazo_determinado", "Fim prazo determinado", False, "data"),
    Posicao(65, "dias_prorrogacao", "Dias da prorrogação do contrato de experiência", False, "digitos"),
    Posicao(66, "prorrogacao_prazo_determinado", "Prorrogação do prazo determinado", False, "data"),
    Posicao(67, "motivo_prorrogacao", "Motivo da prorrogação (até 50 caracteres)", False, "texto"),
    Posicao(68, "forma_calculo_pagamento", "Forma de pagamento (mês/semana)", False, "opcao"),
    Posicao(69, "atividade_simples_nacional", "Atividade para o Simples Nacional", True, "opcao"),
    Posicao(70, "email", "E-mail", False, "email"),
    Posicao(71, "ddd_telefone", "DDD telefone", False, "digitos"),
    Posicao(72, "telefone", "Telefone", False, "digitos"),
    Posicao(73, "ddd_contato_2", "DDD contato 2", False, "digitos"),
    Posicao(74, "contato_2", "Contato 2", False, "digitos"),
    Posicao(75, "", "Data vantagem (= admissão)", True, "data"),
    Posicao(76, "naturalidade_uf", "UF do local de nascimento", True, "uf"),
    Posicao(77, "codigo_categoria_esocial", "Código da categoria eSocial", True, "digitos"),
    Posicao(78, "codigo_sindicato_cadastro", "Código do cadastro do sindicato", True, "digitos"),
]
assert [p.pos for p in POSICOES] == list(range(1, 79))
POSICAO_POR_CHAVE = {p.chave: p for p in POSICOES if p.chave}


# --------------------------------------------------------------------------
# Normalização
# --------------------------------------------------------------------------

def somente_digitos(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def normalizar_decimal(valor: str) -> str | None:
    """"2.500,00" / "2500,5" / "2500" / "7,33" → "2500.00" / "2500.50" / "2500.00" / "7.33"; None se inválido."""
    texto = (valor or "").strip().replace("R$", "").replace(" ", "")
    if not texto:
        return ""
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")  # padrão brasileiro: ponto = milhar
    elif texto.count(".") > 1:
        texto = texto.replace(".", "")
    try:
        numero = Decimal(texto)
    except InvalidOperation:
        return None
    if numero < 0:
        return None
    return f"{numero.quantize(Decimal('0.01'))}"


def parse_data(valor: str) -> date | None:
    texto = (valor or "").strip()
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", texto) or re.fullmatch(r"(\d{2})(\d{2})(\d{4})", texto)
    if not m:
        m2 = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", texto)  # ISO (input type=date)
        if not m2:
            return None
        ano, mes, dia = (int(x) for x in m2.groups())
    else:
        dia, mes, ano = (int(x) for x in m.groups())
    try:
        return date(ano, mes, dia)
    except ValueError:
        return None


def formatar_data(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def mais_um_ano_menos_um_dia(d: date) -> date:
    try:
        um_ano = d.replace(year=d.year + 1)
    except ValueError:  # 29/02 → 01/03 do ano seguinte, como o Excel/VBA fazem
        um_ano = date(d.year + 1, 3, 1)
    return um_ano - timedelta(days=1)


def somar_meses(d: date, meses: int) -> date:
    """Mesmo dia N meses depois (31/01 + 1 mês → 28 ou 29/02)."""
    total = d.month - 1 + meses
    ano, mes = d.year + total // 12, total % 12 + 1
    ultimo_dia = (date(ano + (mes == 12), mes % 12 + 1, 1) - timedelta(days=1)).day
    return date(ano, mes, min(d.day, ultimo_dia))


def fim_do_prazo(inicio: date, dias: int, contagem: str) -> date:
    """Fim de um período de experiência/prorrogação começando em `inicio`.

    CLT (padrão): o 1º dia conta — 30 dias a partir de 01/10 terminam em 30/10.
    MES: como o exemplo da planilha do Domínio — múltiplos de 30 dias viram meses cheios (01/10 + 30 → 01/11).
    """
    if contagem == "MES":
        return somar_meses(inicio, dias // 30) if dias % 30 == 0 else inicio + timedelta(days=dias)
    return inicio + timedelta(days=dias - 1)


def cpf_valido(cpf: str) -> bool:
    d = somente_digitos(cpf)
    if len(d) != 11 or d == d[0] * 11:
        return False
    for tamanho in (9, 10):
        soma = sum(int(d[i]) * (tamanho + 1 - i) for i in range(tamanho))
        digito = (soma * 10 % 11) % 10
        if digito != int(d[tamanho]):
            return False
    return True


def _limpar_texto(valor: str) -> str:
    # TAB ou quebra de linha dentro de um valor deslocaria todas as colunas seguintes no Domínio.
    return re.sub(r"[\t\r\n]+", " ", valor or "").strip()


def _mapear_sexo(valor: str) -> str:
    s = (valor or "").strip().lower()
    if s.startswith("m"):
        return "M"
    if s.startswith("f"):
        return "F"
    return valor.strip()


def _valor_opcao(chave: str, valor: str) -> str:
    """Aceita o código ("S") ou o texto da opção ("Solteiro") e devolve o código."""
    texto = (valor or "").strip()
    if not texto:
        return ""
    opcoes = OPCOES.get(chave, [])
    for codigo, rotulo in opcoes:
        if texto.upper() == codigo.upper():
            return codigo
    alvo = normalizar_nome(texto)
    for codigo, rotulo in opcoes:
        # "Casada"/"Solteira" (como a IA ou o RH escrevem) valem como "Casado"/"Solteiro".
        if alvo in (normalizar_nome(rotulo), re.sub(r"O\b", "A", normalizar_nome(rotulo))):
            return codigo
    if chave.startswith("pais_") and alvo in PAISES:
        return str(PAISES[alvo])
    return texto


# --------------------------------------------------------------------------
# Preparação: normaliza, calcula e confere
# --------------------------------------------------------------------------

@dataclass
class Resultado:
    dados: dict
    calculados: list[str] = field(default_factory=list)  # campos que o sistema preencheu sozinho
    erros: dict[str, str] = field(default_factory=dict)  # impedem gerar o TXT
    avisos: dict[str, str] = field(default_factory=dict)  # só alertam
    dicas: dict[str, str] = field(default_factory=dict)  # explicação do valor que vai pro Domínio


def preparar(entrada: dict) -> Resultado:
    dados = {chave: _limpar_texto(str(valor if valor is not None else "")) for chave, valor in entrada.items()}
    r = Resultado(dados=dados)

    def calcular(chave: str, valor: str) -> None:
        if valor and not dados.get(chave):
            dados[chave] = valor
            r.calculados.append(chave)

    # 1) Normalização por tipo
    dados["sexo"] = _mapear_sexo(dados.get("sexo", ""))
    for p in POSICOES:
        if not p.chave or not dados.get(p.chave):
            continue
        valor = dados[p.chave]
        if p.tipo == "digitos":
            dados[p.chave] = somente_digitos(valor)
        elif p.tipo == "decimal":
            normalizado = normalizar_decimal(valor)
            if normalizado is None:
                r.erros[p.chave] = f"“{valor}” não é um número válido (ex.: 2500.00)."
            else:
                dados[p.chave] = normalizado
        elif p.tipo == "data":
            d = parse_data(valor)
            if d is None:
                r.erros[p.chave] = f"“{valor}” não é uma data válida (DD/MM/AAAA)."
            else:
                dados[p.chave] = formatar_data(d)
        elif p.tipo == "opcao":
            dados[p.chave] = _valor_opcao(p.chave, valor)
        elif p.tipo == "uf":
            dados[p.chave] = valor.strip().upper()

    # 2) Cálculos (só preenchem o que está vazio — o que o RH digitou sempre vale)
    admissao = parse_data(dados.get("data_admissao", ""))
    if admissao:
        calcular("vencimento_ferias", formatar_data(mais_um_ano_menos_um_dia(admissao)))
    calcular("optante_fgts", "S")
    # Padrões do RH (21/10/2026): quase sempre mensalista e admissão normal — o formulário deixa trocar.
    calcular("categoria", "1")
    calcular("tipo_indicativo_admissao", "2")
    calcular("ocorrencia_sefip", "0")
    calcular("multiplos_vinculos", "N")
    calcular("pagou_contribuicao_sindical", "N")
    calcular("deficiente_fisico", "0")
    calcular("tipo_admissao", "1")
    calcular("situacao_qualificacao_cadastral", "0")
    calcular("recebia_seguro_desemprego", "0")
    calcular("forma_calculo_pagamento", "1")
    # Fora do formulário: a filial é a própria empresa em uso no Domínio; banco não é informado.
    calcular("codigo_filial", "1")
    calcular("tipo_conta", "1")
    if dados.get("optante_fgts") == "S" and admissao:
        calcular("data_opcao_fgts", formatar_data(admissao))

    horas = HORAS_POR_SEMANA.get(dados.get("horas_semana", ""))
    if horas:
        calcular("horas_dia", horas[0])
        calcular("horas_mes", horas[1])

    calcular("codigo_categoria_esocial", CATEGORIA_ESOCIAL_POR_VINCULO.get(dados.get("vinculo_empregaticio", ""), ""))

    primeira_palavra = normalizar_nome(dados.get("endereco_logradouro", "")).replace(".", " ").split(" ")[0] if dados.get("endereco_logradouro") else ""
    # Fora do formulário: logradouro sem prefixo conhecido vai como "99 - Outro".
    calcular("tipo_endereco", TIPO_ENDERECO_POR_PREFIXO.get(primeira_palavra, "99" if primeira_palavra else ""))

    calcular("codigo_municipio_endereco", codigo_municipio(dados.get("endereco_cidade", ""), dados.get("endereco_uf", "")))
    calcular("codigo_municipio_nascimento", codigo_municipio(dados.get("naturalidade_cidade", ""), dados.get("naturalidade_uf", "")))
    calcular("pais_endereco", PAIS_BRASIL if dados.get("endereco_uf") in UFS else "")
    calcular("pais_nascimento", PAIS_BRASIL if dados.get("naturalidade_uf") in UFS else "")
    # Estrangeiro: tenta o país pelo texto da nacionalidade do contrato (ex.: "Argentina").
    pais_pelo_texto = str(PAISES.get(normalizar_nome(dados.get("nacionalidade", "")), ""))
    calcular("pais_nascimento", pais_pelo_texto)
    calcular("pais_nacionalidade", PAIS_BRASIL if dados.get("pais_nascimento") == PAIS_BRASIL else pais_pelo_texto)

    if dados.get("contrato_experiencia") == "1":
        if admissao:
            calcular("inicio_prazo_determinado", formatar_data(admissao))
        inicio = parse_data(dados.get("inicio_prazo_determinado", ""))
        dias = dados.get("dias_contrato_experiencia", "")
        contagem = dados.get("contagem_experiencia") or "CLT"
        if inicio and dias.isdigit() and int(dias) > 0:
            calcular("fim_prazo_determinado", formatar_data(fim_do_prazo(inicio, int(dias), contagem)))
        fim = parse_data(dados.get("fim_prazo_determinado", ""))
        dias_prorrogacao = dados.get("dias_prorrogacao", "")
        if fim and dias_prorrogacao.isdigit() and int(dias_prorrogacao) > 0:
            # A prorrogação começa no dia seguinte ao fim (CLT) — no "mês cheio" ela soma a partir do próprio fim, como no exemplo.
            base = fim if contagem == "MES" else fim + timedelta(days=1)
            calcular("prorrogacao_prazo_determinado", formatar_data(fim_do_prazo(base, int(dias_prorrogacao), contagem)))

    # 3) Conferência contra a planilha (obrigatórios, códigos do dicionário, tabelas)
    for p in POSICOES:
        if not p.chave or p.chave in r.erros:
            continue
        valor = dados.get(p.chave, "")
        if not valor:
            if p.obrigatorio:
                r.erros[p.chave] = "Obrigatório no Domínio."
            continue
        if p.tipo == "opcao" and valor not in {codigo for codigo, _ in OPCOES[p.chave]}:
            r.erros[p.chave] = f"“{valor}” não é um código válido do dicionário do Domínio."
        elif p.tipo == "uf" and valor not in UFS:
            r.erros[p.chave] = f"“{valor}” não é uma sigla de UF."
        elif p.tipo == "email" and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", valor):
            r.avisos[p.chave] = "E-mail com formato estranho — confira."

    if dados.get("cpf") and "cpf" not in r.erros and not cpf_valido(dados["cpf"]):
        r.erros["cpf"] = "CPF inválido (dígitos verificadores não conferem)."
    if dados.get("endereco_cep") and "endereco_cep" not in r.erros and len(dados["endereco_cep"]) != 8:
        r.erros["endereco_cep"] = "CEP deve ter 8 dígitos."
    for chave in ("ddd_telefone", "ddd_contato_2"):
        dados[chave] = dados.get(chave, "").lstrip("0")  # "021" (como se disca) → "21"
        if dados.get(chave) and len(dados[chave]) != 2:
            r.erros[chave] = "DDD deve ter 2 dígitos."
    if len(dados.get("motivo_prorrogacao", "")) > 50:
        r.erros["motivo_prorrogacao"] = "Máximo de 50 caracteres no Domínio."

    _conferir_municipio(r, "codigo_municipio_endereco", "endereco_cidade", "endereco_uf", "endereço")
    _conferir_municipio(r, "codigo_municipio_nascimento", "naturalidade_cidade", "naturalidade_uf", "nascimento")

    # Regras condicionais do cabeçalho da planilha ("*IMPORTANTE ...")
    if dados.get("contrato_experiencia") == "1":
        for chave in ("dias_contrato_experiencia", "inicio_prazo_determinado", "fim_prazo_determinado"):
            if not dados.get(chave) and chave not in r.erros:
                r.erros[chave] = "Obrigatório quando há contrato de experiência."
    if dados.get("dias_prorrogacao"):
        for chave in ("prorrogacao_prazo_determinado", "motivo_prorrogacao"):
            if not dados.get(chave) and chave not in r.erros:
                r.erros[chave] = "Obrigatório quando há prorrogação."

    if admissao and (nascimento := parse_data(dados.get("data_nascimento", ""))):
        idade = (admissao - nascimento).days // 365
        if idade < 14:
            r.erros["data_nascimento"] = f"Colaborador teria {idade} anos na admissão — confira as datas."
        elif idade < 18 and dados.get("vinculo_empregaticio") not in ("53", "50"):
            r.avisos["data_nascimento"] = f"Menor de idade na admissão ({idade} anos) — confira o vínculo (aprendiz/estagiário)."

    # Dicas: o que de fato vai pro arquivo
    if dados.get("codigo_municipio_endereco") and "codigo_municipio_endereco" not in r.erros:
        r.dicas["endereco_cidade"] = f"→ Domínio: município {dados['codigo_municipio_endereco']}"
    if dados.get("codigo_municipio_nascimento") and "codigo_municipio_nascimento" not in r.erros:
        r.dicas["naturalidade_cidade"] = f"→ Domínio: município {dados['codigo_municipio_nascimento']}"
    if not dados.get("rg") and dados.get("cpf"):
        r.dicas["rg"] = f"→ RG vazio: o Domínio vai usar o CPF ({dados['cpf']}) como identidade"
    _levar_para_campo_visivel(r)
    return r


# Colunas que não aparecem no formulário: o erro vai para o campo de onde o valor é calculado.
CAMPO_VISIVEL = {
    "tipo_endereco": "endereco_logradouro",
    "codigo_municipio_nascimento": "naturalidade_cidade",
    "pais_nascimento": "nacionalidade",
    "pais_nacionalidade": "nacionalidade",
    "pais_endereco": "endereco_uf",
    "codigo_filial": "codigo_empregado",
    "codigo_banco": "codigo_empregado",
    "tipo_conta": "codigo_empregado",
    "optante_fgts": "data_admissao",
    "data_opcao_fgts": "data_admissao",
    "vencimento_ferias": "data_admissao",
    "horas_dia": "horas_semana",
    "horas_mes": "horas_semana",
    "codigo_categoria_esocial": "vinculo_empregaticio",
}
ROTULO_OCULTO = {
    "tipo_endereco": "Tipo de endereço", "codigo_municipio_nascimento": "Município de nascimento",
    "pais_nascimento": "País de nascimento", "pais_nacionalidade": "País da nacionalidade", "pais_endereco": "País do endereço",
    "codigo_filial": "Código filial", "codigo_banco": "Código banco", "tipo_conta": "Tipo da conta",
    "optante_fgts": "Optante do FGTS", "data_opcao_fgts": "Data de opção do FGTS", "vencimento_ferias": "Vencimento das férias",
    "horas_dia": "Horas por dia (o sistema calcula para 44, 40, 36, 30, 25 ou 20 h por semana)", "horas_mes": "Horas por mês",
    "codigo_categoria_esocial": "Categoria eSocial",
}


def _levar_para_campo_visivel(r: Resultado) -> None:
    for oculto, visivel in CAMPO_VISIVEL.items():
        erro = r.erros.pop(oculto, None)
        if not erro:
            continue
        if erro.startswith("Obrigatório"):
            erro = f"{ROTULO_OCULTO[oculto]} não foi calculado — confira este campo."
        r.erros.setdefault(visivel, erro)


def _conferir_municipio(r: Resultado, chave_codigo: str, chave_cidade: str, chave_uf: str, qual: str) -> None:
    dados = r.dados
    codigo = dados.get(chave_codigo, "")
    if not codigo or chave_codigo in r.erros:
        if not codigo and (dados.get(chave_cidade) or dados.get(chave_uf)):
            r.erros[chave_codigo] = (
                f"Município de {qual} “{dados.get(chave_cidade, '')}/{dados.get(chave_uf, '')}” não encontrado na tabela do "
                "Domínio — confira a grafia da cidade e a UF, ou informe o código."
            )
        return
    uf_do_codigo = next((uf for uf, cidades in MUNICIPIOS.items() if int(codigo) in cidades.values()), None)
    if uf_do_codigo is None:
        r.erros[chave_codigo] = f"Código {codigo} não existe na tabela de municípios do Domínio."
    elif dados.get(chave_uf) and uf_do_codigo != dados[chave_uf]:
        r.erros[chave_codigo] = f"O código {codigo} é de um município de {uf_do_codigo}, mas a UF informada é {dados[chave_uf]}."


VALORES_FIXOS = {1: lambda d: "0", 75: lambda d: d.get("data_admissao", "")}


def linha(dados: dict) -> list[str]:
    """As 78 colunas, na ordem da planilha. Identidade (56) cai para o CPF quando não há RG (CIN)."""
    colunas = []
    for p in POSICOES:
        if p.pos in VALORES_FIXOS:
            colunas.append(VALORES_FIXOS[p.pos](dados))
        elif p.pos == 56:
            colunas.append(dados.get("rg", "").strip() or dados.get("cpf", ""))
        else:
            colunas.append(_limpar_texto(dados.get(p.chave, "")))
    return colunas


def metadados_campos() -> dict:
    """Para o frontend: rótulo/posição/obrigatoriedade/opções de cada campo, direto destas regras."""
    campos = {
        p.chave: {"posicao": p.pos, "rotulo": p.rotulo, "obrigatorio": p.obrigatorio, "opcoes": OPCOES.get(p.chave)}
        for p in POSICOES
        if p.chave
    }
    # Campos com lista de opções que não são coluna do layout (ex.: como contar a experiência).
    for chave, opcoes in OPCOES.items():
        campos.setdefault(chave, {"posicao": None, "rotulo": chave, "obrigatorio": False, "opcoes": opcoes})
    return campos
