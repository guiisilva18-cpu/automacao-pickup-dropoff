"""Cliente mínimo da Open API do Feishu pro bot do resumo diário.

Por que app e não o webhook de "bot customizado": o webhook só manda texto e
cartão -- imagem exige `image_key`, que só se obtém subindo o arquivo pela
API de app (im/v1/images, com tenant_access_token). Permissões do app:
`im:message:send_as_bot` (mandar mensagem) e `im:resource` (subir imagem);
quem recebe precisa estar no escopo de disponibilidade do app.

Variáveis de ambiente:
  FEISHU_APP_ID, FEISHU_APP_SECRET  -- credenciais do app
  FEISHU_DESTINO                     -- e-mail (ou open_id/chat_id) de quem recebe
  FEISHU_DESTINO_TIPO                -- email (padrão) | open_id | user_id | chat_id
  FEISHU_BASE_URL                    -- padrão https://open.feishu.cn
"""
import json
import os

import requests

TIMEOUT = 30


class FeishuErro(RuntimeError):
    pass


def app_configurado() -> bool:
    return all(os.environ.get(v) for v in ("FEISHU_APP_ID", "FEISHU_APP_SECRET", "FEISHU_DESTINO"))


def webhook_configurado() -> bool:
    return bool(os.environ.get("FEISHU_WEBHOOK_URL"))


def modo() -> str | None:
    """'imagem' (app: sobe PNG e manda pra pessoa/grupo) ou 'cartao' (webhook
    do grupo: só texto e cartão, o webhook não aceita imagem). FEISHU_MODO
    força um dos dois quando os dois estão configurados."""
    forcado = (os.environ.get("FEISHU_MODO") or "").strip().lower()
    if forcado == "imagem" and app_configurado():
        return "imagem"
    if forcado == "cartao" and webhook_configurado():
        return "cartao"
    if app_configurado():
        return "imagem"
    if webhook_configurado():
        return "cartao"
    return None


def _base() -> str:
    # `or`: no Actions um secret ausente chega como string vazia, não como variável ausente
    return (os.environ.get("FEISHU_BASE_URL") or "https://open.feishu.cn").rstrip("/")


def _checar(resp: requests.Response, contexto: str) -> dict:
    try:
        corpo = resp.json()
    except ValueError:
        raise FeishuErro(f"{contexto}: resposta não-JSON (HTTP {resp.status_code}): {resp.text[:200]}")
    if resp.status_code >= 400 or corpo.get("code", 0) != 0:
        raise FeishuErro(f"{contexto}: HTTP {resp.status_code}, code={corpo.get('code')}, msg={corpo.get('msg')}")
    return corpo


class Feishu:
    def __init__(self):
        self._token = None
        self.destino = os.environ["FEISHU_DESTINO"]
        self.destino_tipo = os.environ.get("FEISHU_DESTINO_TIPO") or "email"

    def _token_acesso(self) -> str:
        if self._token is None:
            resp = requests.post(
                f"{_base()}/open-apis/auth/v3/tenant_access_token/internal",
                json={"app_id": os.environ["FEISHU_APP_ID"], "app_secret": os.environ["FEISHU_APP_SECRET"]},
                timeout=TIMEOUT,
            )
            self._token = _checar(resp, "tenant_access_token")["tenant_access_token"]
        return self._token

    def _auth(self) -> dict:
        return {"Authorization": f"Bearer {self._token_acesso()}"}

    def subir_imagem(self, png: bytes, nome: str = "resumo.png") -> str:
        resp = requests.post(
            f"{_base()}/open-apis/im/v1/images",
            headers=self._auth(),
            data={"image_type": "message"},
            files={"image": (nome, png, "image/png")},
            timeout=TIMEOUT,
        )
        return _checar(resp, f"upload de imagem '{nome}'")["data"]["image_key"]

    def _enviar(self, msg_type: str, content: dict) -> str:
        resp = requests.post(
            f"{_base()}/open-apis/im/v1/messages",
            params={"receive_id_type": self.destino_tipo},
            headers=self._auth(),
            json={"receive_id": self.destino, "msg_type": msg_type, "content": json.dumps(content, ensure_ascii=False)},
            timeout=TIMEOUT,
        )
        return _checar(resp, f"envio de mensagem ({msg_type})")["data"]["message_id"]

    def enviar_texto(self, texto: str) -> str:
        return self._enviar("text", {"text": texto})

    def enviar_imagem(self, image_key: str) -> str:
        return self._enviar("image", {"image_key": image_key})


class Webhook:
    """Bot customizado de grupo (webhook). Só texto e cartão -- imagem não
    passa por aqui. Segurança por palavra-chave: TODA mensagem precisa conter
    FEISHU_KEYWORD (o rodapé de cada cartão/texto leva a palavra). Se o bot
    tiver "Assinatura" ligada em vez de palavra-chave, FEISHU_WEBHOOK_SECRET
    assina a requisição."""

    def __init__(self, url: str | None = None, palavra: str | None = None, segredo: str | None = None):
        """Sem argumentos, usa FEISHU_WEBHOOK_URL/FEISHU_KEYWORD/
        FEISHU_WEBHOOK_SECRET (bot do resumo da manhã, comportamento
        original). Passando os argumentos, manda pra outro webhook/grupo --
        usado pelo bot_poc_21h.py (resumo de "tentativa fora do prazo" às
        21h, webhook diferente do da manhã, 07/10/2026)."""
        self.url = (url if url is not None else os.environ["FEISHU_WEBHOOK_URL"]).strip()
        self.palavra = (palavra if palavra is not None else (os.environ.get("FEISHU_KEYWORD") or "")).strip()
        self.segredo = (segredo if segredo is not None else (os.environ.get("FEISHU_WEBHOOK_SECRET") or "")).strip()

    @property
    def rodape(self) -> str:
        return f"Resumo automático · {self.palavra}" if self.palavra else "Resumo automático"

    def _assinar(self, payload: dict) -> dict:
        if self.segredo:
            import base64
            import hashlib
            import hmac
            import time

            ts = str(int(time.time()))
            chave = f"{ts}\n{self.segredo}".encode("utf-8")
            payload["timestamp"] = ts
            payload["sign"] = base64.b64encode(hmac.new(chave, digestmod=hashlib.sha256).digest()).decode("utf-8")
        return payload

    def _enviar(self, payload: dict, contexto: str):
        resp = requests.post(self.url, json=self._assinar(payload), timeout=TIMEOUT)
        try:
            corpo = resp.json()
        except ValueError:
            raise FeishuErro(f"{contexto}: resposta não-JSON (HTTP {resp.status_code}): {resp.text[:200]}")
        codigo = corpo.get("code", corpo.get("StatusCode", 0))
        if resp.status_code >= 400 or codigo != 0:
            raise FeishuErro(f"{contexto}: HTTP {resp.status_code}, code={codigo}, msg={corpo.get('msg') or corpo.get('StatusMessage')}")

    def enviar_texto(self, texto: str):
        if self.palavra and self.palavra not in texto:
            texto = f"{texto}\n{self.rodape}"
        self._enviar({"msg_type": "text", "content": {"text": texto}}, "texto")

    def enviar_cartao(self, cartao: dict):
        self._enviar({"msg_type": "interactive", "card": cartao}, "cartão")
