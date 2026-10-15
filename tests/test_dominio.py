"""Testes do layout do Domínio contra a planilha oficial (data/fonte/empregados_modelo_dominio.xls).

Roda sem pytest: `python tests/test_dominio.py` (também funciona com pytest).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import dominio  # noqa: E402

# Linha da aba "Exemplo" da planilha, como a macro "Gerar" grava (datas em DD/MM/AAAA).
EXEMPLO_PLANILHA = [
    "0", "1", "Renan Massirolli Exemplo", "12345678909", "1", "1", "3", "1", "1", "8000.00",
    "01/10/2024", "77", "1", "S", "01/10/2024", "1", "RUA PROJETADA", "1000", "PORTÃO GRADEADO", "CENTRO",
    "CRICIUMA", "SC", "88860000", "01/07/1990", "1", "123456789101213", "1", "2", "4", "N",
    "S", "30/09/2025", "7.33", "1", "2", "9", "0", "1", "44.00", "C",
    "220.00", "0", "M", "30", "2", "S", "0", "1", "Walme Massirolli", "Rosimeri Machado",
    "33", "4489", "4489", "30", "30", "1.23.456", "SSP", "SC", "01/09/2000", "1",
    "1", "30", "01/10/2024", "01/11/2024", "30", "01/12/2024", "Outros", "1", "3", "email@exemplo.com.br",
    "48", "991690001", "48", "991690002", "01/10/2024", "SC", "101", "1",
]

# O mesmo empregado como o RH preencheria no sistema (formatos "humanos"; o resto é calculado).
ENTRADA_EXEMPLO = {
    "codigo_empregado": "1", "nome_completo": "Renan Massirolli Exemplo", "cpf": "123.456.789-09",
    "codigo_filial": "1", "codigo_servico": "1", "codigo_cargo": "3", "codigo_departamento": "1",
    "codigo_centro_custos": "1", "salario": "8.000,00", "data_admissao": "01/10/2024",
    "codigo_esocial_empregado": "77", "vinculo_empregaticio": "1", "ocorrencia_sefip": "1",
    "endereco_logradouro": "RUA PROJETADA", "endereco_numero": "1000", "endereco_complemento": "PORTÃO GRADEADO",
    "endereco_bairro": "CENTRO", "endereco_cidade": "CRICIUMA", "endereco_uf": "sc", "endereco_cep": "88860-000",
    "data_nascimento": "01/07/1990", "codigo_banco": "1", "conta_corrente": "123456789101213", "digito_conta": "1",
    "tipo_conta": "2", "codigo_convenio_coletivo": "4", "multiplos_vinculos": "N", "pagou_contribuicao_sindical": "S",
    "categoria": "1", "tipo_indicativo_admissao": "2", "grau_instrucao": "9", "deficiente_fisico": "0",
    "tipo_admissao": "1", "horas_semana": "44", "forma_pagamento": "C", "situacao_qualificacao_cadastral": "0",
    "sexo": "Masculino", "raca_cor": "2", "estado_civil": "Solteiro", "recebia_seguro_desemprego": "0",
    "codigo_jornada": "1", "nome_pai": "Walme Massirolli", "nome_mae": "Rosimeri Machado",
    # Naturalidade com acento/caixa diferentes do endereço: a busca na tabela acha o mesmo 4489.
    "naturalidade_cidade": "Criciúma", "naturalidade_uf": "SC", "rg": "1.23.456", "rg_orgao_emissor": "SSP",
    "rg_uf_emissor": "SC", "rg_data_emissao": "01/09/2000", "contrato_experiencia": "1",
    "clausula_assecuratoria": "1", "dias_contrato_experiencia": "30",
    # O exemplo da planilha usa fim = 01/11 e prorrogação = 01/12 (mês cheio); o sistema calcularia pela
    # contagem da CLT (30/10 e 29/11) — aqui o RH informa as datas, como no exemplo.
    "fim_prazo_determinado": "01/11/2024", "dias_prorrogacao": "30", "prorrogacao_prazo_determinado": "01/12/2024",
    "motivo_prorrogacao": "Outros", "forma_calculo_pagamento": "1", "atividade_simples_nacional": "3",
    "email": "email@exemplo.com.br", "ddd_telefone": "(48)", "telefone": "99169-0001", "ddd_contato_2": "48",
    "contato_2": "99169-0002", "codigo_sindicato_cadastro": "1",
}


def test_exemplo_da_planilha_bate_coluna_a_coluna():
    r = dominio.preparar(ENTRADA_EXEMPLO)
    assert r.erros == {}, r.erros
    obtido = dominio.linha(r.dados)
    assert len(obtido) == 78
    diferencas = [(i + 1, esperado, valor) for i, (esperado, valor) in enumerate(zip(EXEMPLO_PLANILHA, obtido)) if esperado != valor]
    assert not diferencas, diferencas


def test_calculados_do_exemplo():
    r = dominio.preparar(ENTRADA_EXEMPLO)
    for chave in ("vencimento_ferias", "optante_fgts", "data_opcao_fgts", "horas_dia", "horas_mes",
                  "codigo_categoria_esocial", "tipo_endereco", "codigo_municipio_endereco",
                  "codigo_municipio_nascimento", "pais_endereco", "pais_nascimento", "pais_nacionalidade",
                  "inicio_prazo_determinado"):
        assert chave in r.calculados, chave


def test_decimal_formato_brasileiro():
    assert dominio.normalizar_decimal("2.500,00") == "2500.00"
    assert dominio.normalizar_decimal("2500,5") == "2500.50"
    assert dominio.normalizar_decimal("R$ 1.412,00") == "1412.00"
    assert dominio.normalizar_decimal("8000.00") == "8000.00"
    assert dominio.normalizar_decimal("7,33") == "7.33"
    assert dominio.normalizar_decimal("abc") is None


def test_municipio_homonimo_respeita_uf():
    assert dominio.codigo_municipio("Rio Branco", "AC") == "16"
    assert dominio.codigo_municipio("Rio Branco", "MT") == "2384"
    assert dominio.codigo_municipio("Niterói", "RJ") == "3638"
    assert dominio.codigo_municipio("Niterói", "SP") == ""


def test_municipio_de_outra_uf_e_cidade_inexistente_dao_erro():
    r = dominio.preparar({"endereco_cidade": "Niterói", "endereco_uf": "RJ", "codigo_municipio_endereco": "16"})
    assert "AC" in r.erros["codigo_municipio_endereco"]
    r = dominio.preparar({"endereco_cidade": "Niteroi do Norte", "endereco_uf": "RJ"})
    assert "não encontrado" in r.erros["codigo_municipio_endereco"]


def test_experiencia_pela_clt_e_condicionais():
    r = dominio.preparar({"data_admissao": "01/10/2024", "contrato_experiencia": "1", "dias_contrato_experiencia": "45",
                          "dias_prorrogacao": "45"})
    assert r.dados["inicio_prazo_determinado"] == "01/10/2024"
    assert r.dados["fim_prazo_determinado"] == "14/11/2024"
    assert r.dados["prorrogacao_prazo_determinado"] == "29/12/2024"
    assert r.erros["motivo_prorrogacao"] == "Obrigatório quando há prorrogação."


def test_contagem_mes_cheio_bate_com_o_exemplo_da_planilha():
    r = dominio.preparar({"data_admissao": "01/10/2024", "contrato_experiencia": "1", "contagem_experiencia": "MES",
                          "dias_contrato_experiencia": "30", "dias_prorrogacao": "30"})
    assert r.dados["fim_prazo_determinado"] == "01/11/2024"
    assert r.dados["prorrogacao_prazo_determinado"] == "01/12/2024"
    r = dominio.preparar({"data_admissao": "01/10/2024", "contrato_experiencia": "1", "contagem_experiencia": "CLT",
                          "dias_contrato_experiencia": "30", "dias_prorrogacao": "30"})
    assert r.dados["fim_prazo_determinado"] == "30/10/2024"
    assert r.dados["prorrogacao_prazo_determinado"] == "29/11/2024"
    assert dominio.somar_meses(dominio.date(2024, 1, 31), 1) == dominio.date(2024, 2, 29)


def test_validacoes_basicas():
    r = dominio.preparar({"cpf": "111.111.111-11", "endereco_cep": "2422", "sexo": "X", "data_admissao": "31/02/2024",
                          "ddd_telefone": "2", "estado_civil": "Casada"})
    assert "CPF inválido" in r.erros["cpf"]
    assert "8 dígitos" in r.erros["endereco_cep"]
    assert "não é um código válido" in r.erros["sexo"]
    assert "data válida" in r.erros["data_admissao"]
    assert "2 dígitos" in r.erros["ddd_telefone"]
    assert r.dados["estado_civil"] == "C"  # “Casada” vale como “Casado”


def test_tab_dentro_do_valor_nao_quebra_colunas():
    r = dominio.preparar({**ENTRADA_EXEMPLO, "endereco_complemento": "BLOCO\t2\nFUNDOS"})
    obtido = dominio.linha(r.dados)
    assert obtido[18] == "BLOCO 2 FUNDOS"
    assert len("\t".join(obtido).split("\t")) == 78


def test_ferias_em_29_de_fevereiro():
    r = dominio.preparar({"data_admissao": "29/02/2024"})
    assert r.dados["vencimento_ferias"] == "28/02/2025"


def test_rh_sempre_vence_o_calculo():
    r = dominio.preparar({**ENTRADA_EXEMPLO, "vencimento_ferias": "15/10/2025", "tipo_endereco": "4"})
    assert r.dados["vencimento_ferias"] == "15/10/2025"
    assert r.dados["tipo_endereco"] == "4"
    assert "vencimento_ferias" not in r.calculados


def test_campos_do_layout_existem_no_modelo():
    from models import DadosAdmissao

    campos = set(DadosAdmissao.model_fields)
    faltando = [p.chave for p in dominio.POSICOES if p.chave and p.chave not in campos]
    assert not faltando, faltando


if __name__ == "__main__":
    falhas = 0
    for nome, teste in list(globals().items()):
        if nome.startswith("test_") and callable(teste):
            try:
                teste()
                print(f"ok    {nome}")
            except AssertionError as exc:
                falhas += 1
                print(f"FALHA {nome}: {exc}")
    sys.exit(1 if falhas else 0)
