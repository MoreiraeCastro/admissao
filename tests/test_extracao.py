"""Extração pela IA: formato fixo, segunda leitura só dos campos recusados e fonte/dúvida por campo (cliente falso, sem API)."""
import json
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import anthropic  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

import main  # noqa: E402


def _resposta(dados, fontes=(), duvidas=()):
    corpo = {"dados": {c: dados.get(c, "") for c in main.CAMPOS_IA}, "fontes": list(fontes), "duvidas": list(duvidas)}
    return SimpleNamespace(stop_reason="end_turn", content=[SimpleNamespace(type="text", text=json.dumps(corpo))])


class ClienteFalso:
    def __init__(self, respostas):
        self.respostas, self.pedidos = list(respostas), []
        self.messages = SimpleNamespace(create=self._create)

    def _create(self, **kwargs):
        self.pedidos.append(kwargs)
        return self.respostas.pop(0)


def _extrair(monkeypatch, respostas):
    cliente = ClienteFalso(respostas)
    monkeypatch.setenv("ANTHROPIC_API_KEY", "teste")
    monkeypatch.setattr(anthropic, "Anthropic", lambda **_: cliente)
    r = TestClient(main.app).post("/api/extrair", files=[("documentos", ("rg.png", b"\x89PNG fake", "image/png"))])
    return r, cliente


def test_cpf_recusado_dispara_segunda_leitura_so_dele(monkeypatch):
    primeira = _resposta({"nome_completo": "Maria Silva", "cpf": "123.456.789-00", "numero_ctps": "0012345"},
                         fontes=[{"campo": "cpf", "documento": "rg.png"}, {"campo": "numero_ctps", "documento": "rg.png"}])
    segunda = _resposta({"nome_completo": "OUTRO NOME", "cpf": "123.456.789-09"}, duvidas=[{"campo": "cpf", "motivo": "último dígito borrado"}])
    r, cliente = _extrair(monkeypatch, [primeira, segunda])
    assert r.status_code == 200, r.text
    corpo = r.json()
    assert len(cliente.pedidos) == 2
    assert "cpf" in cliente.pedidos[1]["messages"][0]["content"][-1]["text"]
    assert corpo["dados"]["cpf"] == "12345678909" and "cpf" not in corpo["erros"]
    assert corpo["dados"]["nome_completo"] == "Maria Silva"  # a segunda leitura só troca os campos recusados
    assert corpo["dados"]["numero_ctps"] == "0012345"  # zeros à esquerda preservados
    assert corpo["ia"] == {"fontes": {"cpf": "rg.png", "numero_ctps": "rg.png"}, "duvidas": {"cpf": "último dígito borrado"}, "relidos": ["cpf"]}
    pedido = cliente.pedidos[0]
    assert pedido["extra_body"]["output_config"]["format"]["schema"] is main.SCHEMA_EXTRACAO
    assert pedido["messages"][0]["content"][-2]["cache_control"] == {"type": "ephemeral"}


def test_sem_erro_de_leitura_faz_uma_chamada_so(monkeypatch):
    r, cliente = _extrair(monkeypatch, [_resposta({"nome_completo": "Maria Silva", "cpf": "123.456.789-09"})])
    assert r.status_code == 200 and len(cliente.pedidos) == 1


def test_recusa_da_ia_vira_mensagem_clara(monkeypatch):
    recusa = SimpleNamespace(stop_reason="refusal", content=[])
    r, _ = _extrair(monkeypatch, [recusa])
    assert r.status_code == 502 and "recusou" in r.json()["detail"]
