# Admissão RH

Sistema interno da Moreira & Castro Assessoria Contábil para automatizar a admissão de
funcionários: leitura de documentos via IA, validação pelo RH e geração dos arquivos
finais (contrato + importação no Domínio Web).

Em produção em `https://portalmoreiraecastro.com.br/admissao/`.

## Como funciona

1. O RH arrasta todos os documentos do colaborador no frontend (`static/index.html`), sem
   limite de quantidade: RG, CPF, CNH, comprovante de residência etc. As fotos são reduzidas
   no navegador antes do envio; o limite é de ~19 MB por envio.
2. `POST /api/extrair` envia os documentos para a API da Anthropic (Claude), que lê e
   devolve os dados pessoais e de endereço em JSON.
3. O RH revisa os dados extraídos e preenche manualmente os campos de folha de
   pagamento (salário, admissão, FGTS, banco, jornada etc. — tudo que o layout do
   Domínio Web exige).
4. Ao aprovar, `POST /api/gerar` gera três arquivos, ligados a um `job_id` (sem banco
   de dados — tudo fica em memória e em disco até a limpeza diária):
   - `contrato_trabalho.docx` — a partir de `template_contrato.docx`, substituindo
     placeholders `{{campo}}`
   - `layout_dominio.txt` — 78 colunas separadas por TAB, `cp1252`, no layout oficial
     de importação de admissão do Domínio Web (só sai quando passa na conferência — ver abaixo)
   - `conferencia_dominio.txt` — versão legível dos mesmos 78 campos, só para o RH
     conferir antes de importar (não é o arquivo que vai pro Domínio)
5. Os três ficam disponíveis para download em `GET /api/download/{job_id}/{tipo}` até
   a limpeza automática, todo dia às 3h (horário de Brasília).

## Stack

- Backend: FastAPI (Python 3.12.4), sem banco de dados
- Frontend: HTML/JS puro (`static/index.html`), sem framework/build step
- IA: API da Anthropic (`anthropic` SDK), modelo configurável via `ANTHROPIC_MODEL`
  (default `claude-sonnet-5-5`)
- Geração de `.docx`: `python-docx`

## Rodando localmente

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
# ou: source .venv/bin/activate  # Linux/Mac

pip install -r requirements.txt
copy .env.example .env           # Windows; ou: cp .env.example .env
# edite o .env e preencha ANTHROPIC_API_KEY

uvicorn main:app --reload
```

Acesse `http://localhost:8000`.

### Variáveis de ambiente

| Variável            | Obrigatória | Descrição                                                        |
|---------------------|:-----------:|--------------------------------------------------------------------|
| `ANTHROPIC_API_KEY` | Sim         | Chave da API da Anthropic, usada em `/api/extrair`                 |
| `ANTHROPIC_MODEL`   | Não         | Modelo Claude a usar (default: `claude-sonnet-5-5`)                  |

## Deploy (produção — AWS EC2)

O deploy é manual, direto num EC2 compartilhado com os outros sistemas do escritório
(FiscalMC/emissor, certificados), atrás do mesmo nginx.

- App roda como serviço systemd `admissao-rh.service`
  (`uvicorn main:app --host 127.0.0.1 --port 8000`, dentro de `.venv`)
- nginx expõe `portalmoreiraecastro.com.br/admissao/` fazendo proxy para
  `127.0.0.1:8000`

Para atualizar produção:

```bash
ssh ubuntu@<ip-do-servidor>
cd /home/ubuntu/admissao-rh
git pull
# se requirements.txt mudou: .venv/bin/pip install -r requirements.txt
sudo systemctl restart admissao-rh
```

Não há CI/CD — o repositório neste EC2 é atualizado manualmente via `git pull`.

## Estrutura do projeto

```
main.py                        # FastAPI: extração via IA, geração de arquivos, download, limpeza diária
dominio.py                     # REGRAS do layout do Domínio: 78 posições, obrigatórios, códigos do
                               # dicionário, normalização, cálculos automáticos e conferência
models.py                      # DadosAdmissao (Pydantic) — todos os campos do fluxo
data/fonte/empregados_modelo_dominio.xls  # planilha oficial do Domínio (com a macro "Gerar") — fonte de tudo
data/municipios_dominio.json   # {UF: {MUNICÍPIO: código}} — aba "Código dos municípios", separada por UF
data/paises_dominio.json       # {PAÍS: código} — aba "Código dos países" (Brasil = 30)
template_contrato.docx         # modelo real do contrato, com placeholders {{campo}}
static/index.html              # frontend (upload, conferência ao vivo, download)
tests/test_dominio.py          # testes do layout (python tests/test_dominio.py)
scripts/create_template.py     # gera um template_contrato.docx MOCK, só para testes
scripts/preparar_template_real.py  # script usado uma vez para converter o contrato
                                    # original da empresa nos placeholders {{campo}}
```

## O TXT do Domínio

O `layout_dominio.txt` substitui a macro "Gerar" da planilha oficial: as mesmas 78
colunas por TAB, sem cabeçalho, ANSI (cp1252) e CRLF. O teste
`test_exemplo_da_planilha_bate_coluna_a_coluna` garante que a linha da aba "Exemplo"
sai idêntica, coluna por coluna.

Antes de gerar, `dominio.preparar()` confere tudo contra a planilha e **bloqueia** o
TXT enquanto houver erro:

- obrigatórios (`*` da aba "Preencher") e as regras condicionais do cabeçalho
  (experiência → dias/início/fim; prorrogação → data/motivo);
- códigos do "Dicionário dos campos" (aceita o código ou o texto, ex. "Casada" → C);
- município pela cidade **e UF** (242 nomes se repetem entre estados) e país pela aba
  de países;
- formatos: ponto decimal (`2.500,00` → `2500.00`), datas DD/MM/AAAA, CPF com dígito
  verificador, CEP com 8 dígitos, DDD com 2, TAB/quebra de linha removidos dos textos.

Calcula sozinho o que dá (sem sobrescrever o que o RH digitou): vencimento de férias,
FGTS (S + data = admissão), horas dia/mês pelas horas semanais, categoria eSocial pelo
vínculo, tipo de endereço pelo logradouro, códigos de município/país e as datas do
contrato de experiência (contagem da CLT: 30 dias a partir de 01/10 terminam em 30/10).

### Atenção ao editar campos

As regras de cada posição ficam em `dominio.py` (o formulário busca opções,
obrigatoriedade e posição em `GET /api/campos`). Um campo novo precisa existir em
`DadosAdmissao` (`models.py`) e em `FIELD_GROUPS` (`static/index.html`) — o teste
`test_campos_do_layout_existem_no_modelo` acusa quando o layout aponta pra um campo
que não existe.
