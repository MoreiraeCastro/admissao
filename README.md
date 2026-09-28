# Admissão RH

Sistema interno da Moreira & Castro Assessoria Contábil para automatizar a admissão de
funcionários: leitura de documentos via IA, validação pelo RH e geração dos arquivos
finais (contrato + importação no Domínio Web).

Em produção em `https://portalmoreiraecastro.com.br/admissao/`.

## Como funciona

1. O RH arrasta os documentos do colaborador (RG, CPF ou CNH e, opcionalmente, um
   comprovante de residência) no frontend (`static/index.html`).
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
     de importação de admissão do Domínio Web
   - `conferencia_dominio.txt` — versão legível dos mesmos 78 campos, só para o RH
     conferir antes de importar (não é o arquivo que vai pro Domínio)
5. Os três ficam disponíveis para download em `GET /api/download/{job_id}/{tipo}` até
   a limpeza automática, todo dia às 3h (horário de Brasília).

## Stack

- Backend: FastAPI (Python 3.12.4), sem banco de dados
- Frontend: HTML/JS puro (`static/index.html`), sem framework/build step
- IA: API da Anthropic (`anthropic` SDK), modelo configurável via `ANTHROPIC_MODEL`
  (default `claude-sonnet-5`)
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
| `ANTHROPIC_MODEL`   | Não         | Modelo Claude a usar (default: `claude-sonnet-5`)                  |

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
models.py                      # DadosAdmissao (Pydantic) — todos os campos do fluxo
data/municipios_dominio.json   # nome do município -> código oficial do Domínio Web
template_contrato.docx         # modelo real do contrato, com placeholders {{campo}}
static/index.html              # frontend (upload, validação, download)
scripts/create_template.py     # gera um template_contrato.docx MOCK, só para testes
scripts/preparar_template_real.py  # script usado uma vez para converter o contrato
                                    # original da empresa nos placeholders {{campo}}
```

### Atenção ao editar campos

Os campos de `DadosAdmissao` (`models.py`), do formulário (`FIELD_GROUPS` em
`static/index.html`) e do mapeamento de posições do layout do Domínio
(`_montar_linha_dominio` em `main.py`) precisam ficar sempre sincronizados — adicionar
ou renomear um campo em um lugar sem replicar nos outros dois quebra o fluxo.
