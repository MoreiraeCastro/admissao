from pydantic import BaseModel

# Campos que a IA procura nos documentos. Usado para filtrar a resposta do prompt de
# extração — o resto de DadosAdmissao é preenchido pelo RH ou calculado (dominio.py).
CAMPOS_ADMISSAO = [
    # Identificação (RG, CPF, CNH, CIN)
    "nome_completo",
    "cpf",
    "rg",
    "rg_orgao_emissor",
    "rg_uf_emissor",
    "rg_data_emissao",
    "data_nascimento",
    "sexo",
    "naturalidade_cidade",
    "naturalidade_uf",
    "nacionalidade",
    "pais_nascimento",
    "nome_mae",
    "nome_pai",
    # Comprovante de residência
    "endereco_logradouro",
    "endereco_numero",
    "endereco_complemento",
    "endereco_bairro",
    "endereco_cidade",
    "endereco_uf",
    "endereco_cep",
    # Outros documentos, quando enviados
    "estado_civil",  # certidão de casamento / averbação
    "grau_instrucao",  # diploma, certificado, histórico
    "numero_ctps",  # CTPS (contrato)
    "serie_ctps",
    "conta_corrente",  # cartão/extrato/comprovante bancário
    "digito_conta",
    "tipo_conta",
    "deficiente_fisico",  # laudo PcD
    "email",  # ficha cadastral, formulário, conta
    "ddd_telefone",
    "telefone",
    "ddd_contato_2",
    "contato_2",
]


class DadosAdmissao(BaseModel):
    # --- Documentais (extraídos pela IA) ---
    nome_completo: str = ""
    cpf: str = ""
    rg: str = ""
    rg_orgao_emissor: str = ""
    rg_uf_emissor: str = ""
    rg_data_emissao: str = ""
    data_nascimento: str = ""
    sexo: str = ""
    naturalidade_cidade: str = ""
    naturalidade_uf: str = ""
    nome_mae: str = ""
    nome_pai: str = ""
    endereco_logradouro: str = ""
    endereco_numero: str = ""
    endereco_complemento: str = ""
    endereco_bairro: str = ""
    endereco_cidade: str = ""
    endereco_uf: str = ""
    endereco_cep: str = ""

    # --- Folha de pagamento (preenchidos manualmente pelo RH) ---
    codigo_empregado: str = ""
    codigo_filial: str = ""
    codigo_servico: str = ""
    codigo_cargo: str = ""
    codigo_departamento: str = ""
    codigo_centro_custos: str = ""
    codigo_convenio_coletivo: str = ""
    codigo_sindicato_cadastro: str = ""

    salario: str = ""
    data_admissao: str = ""
    codigo_esocial_empregado: str = ""
    vinculo_empregaticio: str = ""
    categoria: str = ""
    tipo_indicativo_admissao: str = ""
    tipo_admissao: str = ""
    codigo_categoria_esocial: str = ""
    atividade_simples_nacional: str = ""

    horas_dia: str = ""
    horas_semana: str = ""
    horas_mes: str = ""
    codigo_jornada: str = ""
    forma_calculo_pagamento: str = ""
    forma_pagamento: str = ""
    vencimento_ferias: str = ""

    optante_fgts: str = ""
    data_opcao_fgts: str = ""
    ocorrencia_sefip: str = ""
    multiplos_vinculos: str = ""
    pagou_contribuicao_sindical: str = ""

    codigo_banco: str = ""
    conta_corrente: str = ""
    digito_conta: str = ""
    tipo_conta: str = ""

    contrato_experiencia: str = ""
    clausula_assecuratoria: str = ""
    dias_contrato_experiencia: str = ""
    inicio_prazo_determinado: str = ""
    fim_prazo_determinado: str = ""
    dias_prorrogacao: str = ""
    prorrogacao_prazo_determinado: str = ""
    motivo_prorrogacao: str = ""

    grau_instrucao: str = ""
    deficiente_fisico: str = ""
    situacao_qualificacao_cadastral: str = ""
    raca_cor: str = ""
    estado_civil: str = ""
    recebia_seguro_desemprego: str = ""
    tipo_endereco: str = ""

    # --- Códigos de tabela do Domínio (calculados a partir de cidade/UF/país; o RH pode corrigir) ---
    codigo_municipio_endereco: str = ""
    codigo_municipio_nascimento: str = ""
    pais_nacionalidade: str = ""
    pais_endereco: str = ""
    pais_nascimento: str = ""

    email: str = ""
    ddd_telefone: str = ""
    telefone: str = ""
    ddd_contato_2: str = ""
    contato_2: str = ""

    # --- Só para o texto do contrato (não vão pro arquivo do Domínio) ---
    nacionalidade: str = ""
    numero_ctps: str = ""
    serie_ctps: str = ""
    funcao: str = ""
    desconto_alimentacao: str = ""
    horario_seg_qui: str = ""
    horario_sexta: str = ""
    representante_legal: str = ""


class GerarArquivosResponse(BaseModel):
    job_id: str
    docx_url: str
    dominio_url: str
    conferencia_url: str
