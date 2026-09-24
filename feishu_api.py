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


def configurado() -> bool:
    return all(os.environ.get(v) for v in ("FEISHU_APP_ID", "FEISHU_APP_SECRET", "FEISHU_DESTINO"))


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
