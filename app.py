#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Webhook do WhatsApp Business Cloud API (Meta) -> repassa cada mensagem
recebida (texto + anexos: foto, PDF, áudio) como um e-mail pra
ocarmelin@ocarmelin.com.br, reaproveitando o mesmo fluxo de captura de
dados que já lê e-mails de outros remetentes (Vitru, Santa Casa, Marimed).

Hospedado no Render.com (grátis pro volume esperado). Precisa das
variáveis de ambiente abaixo configuradas no painel do Render (nunca
digitadas direto no código):

  VERIFY_TOKEN     - qualquer string que você inventar (usada só na hora
                      de configurar o webhook lá na Meta, pra provar que
                      é você mesmo configurando).
  WHATSAPP_TOKEN    - o token de acesso PERMANENTE do app da Meta (não o
                      temporário da Etapa 1 — esse expira em horas).

O envio do e-mail de repasse é feito via Microsoft Graph (API-only, sem
SMTP — o Render bloqueia conexões SMTP de saída, por isso NÃO usamos
smtplib aqui):

  MS_CLIENT_ID      - Application (client) ID do app "robo-whatsapp-email"
                      no Azure AD.
  MS_CLIENT_SECRET  - client secret desse mesmo app.
  MS_TENANT_ID      - Directory (tenant) ID da conta Microsoft 365.
  REMETENTE_EMAIL   - ocarmelin@ocarmelin.com.br (caixa usada tanto como
                      remetente quanto destinatária do e-mail de repasse;
                      o app precisa da permissão de aplicativo Mail.Send
                      já concedida/consentida no Azure).

NOVO (01/10/2026, a pedido) — também expõe POST /responder: o robô de
emissão (rodando no PC, depois de emitir uma nota) chama esse endpoint
pra devolver a confirmação (link + chave) pro WhatsApp de quem pediu.
Variáveis de ambiente adicionais pra isso:

  WHATSAPP_PHONE_NUMBER_ID - o "Phone number ID" do número 3225-4911 no
                      Business Manager da Meta (não é o número de
                      telefone em si — é um ID numérico interno).
  RESPONDER_SECRET  - outra senha qualquer que você inventar (diferente
                      do VERIFY_TOKEN) — o robô manda ela num cabeçalho
                      pra provar que é ele mesmo chamando, não qualquer
                      um na internet.
"""

import base64
import os
import time

import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

VERIFY_TOKEN = os.environ["VERIFY_TOKEN"]
WHATSAPP_TOKEN = os.environ["WHATSAPP_TOKEN"]

# Envio de e-mail via Microsoft Graph (não SMTP — ver docstring acima).
MS_CLIENT_ID = os.environ["MS_CLIENT_ID"]
MS_CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
MS_TENANT_ID = os.environ["MS_TENANT_ID"]
REMETENTE_EMAIL = os.environ["REMETENTE_EMAIL"]
DESTINO_EMAIL = os.environ.get("DESTINO_EMAIL", REMETENTE_EMAIL)

# NOVO — usados só pelo endpoint /responder (ver docstring acima). Lidos
# com .get() (não os["..."]) pra não derrubar o serviço inteiro se ainda
# não tiverem sido configurados — nesse caso /responder só devolve erro
# 503 explicando o que falta, o resto do webhook (receber mensagem) segue
# funcionando normalmente.
WHATSAPP_PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID", "")
RESPONDER_SECRET = os.environ.get("RESPONDER_SECRET", "")

GRAPH_URL = "https://graph.facebook.com/v25.0"

# --- cache simples do token do Graph (client-credentials), pra não pedir
# um token novo a cada e-mail — eles duram ~1h, renovamos uns minutos antes
# de expirar.
_graph_token_cache = {"token": None, "expira_em": 0}


def obter_token_graph():
    agora = time.time()
    if _graph_token_cache["token"] and agora < _graph_token_cache["expira_em"] - 120:
        return _graph_token_cache["token"]

    resp = requests.post(
        f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=20,
    )
    resp.raise_for_status()
    dados = resp.json()
    _graph_token_cache["token"] = dados["access_token"]
    _graph_token_cache["expira_em"] = agora + int(dados.get("expires_in", 3600))
    return _graph_token_cache["token"]


@app.route("/webhook", methods=["GET"])
def verificar_webhook():
    """A Meta chama isso UMA VEZ, na hora de você configurar a URL de
    callback lá no painel — precisa devolver exatamente o "hub.challenge"
    que ela manda, senão a verificação falha."""
    modo = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if modo == "subscribe" and token == VERIFY_TOKEN:
        return challenge, 200
    return "Verificação falhou — token não confere.", 403


@app.route("/webhook", methods=["POST"])
def receber_webhook():
    """A Meta chama isso toda vez que chega mensagem nova no número
    cadastrado. Sempre respondemos 200 rápido (mesmo se algo dentro der
    erro) — senão a Meta interpreta como falha e fica reenviando."""
    dados = request.get_json(silent=True) or {}
    try:
        for entrada in dados.get("entry", []):
            for mudanca in entrada.get("changes", []):
                valor = mudanca.get("value", {})
                contatos = {
                    c["wa_id"]: c.get("profile", {}).get("name", "")
                    for c in valor.get("contacts", [])
                }
                for msg in valor.get("messages", []):
                    processar_mensagem(msg, contatos)
    except Exception as exc:  # noqa: BLE001 — nunca pode derrubar o endpoint
        print(f"[ERRO] processando webhook: {exc}")
    return jsonify({"status": "ok"}), 200


@app.route("/")
def raiz():
    return "OK — webhook do robô de captura WhatsApp está no ar."


@app.route("/responder", methods=["POST"])
def responder():
    """NOVO (01/10/2026, a pedido) — o robô de emissão (no PC) chama isso
    depois de emitir uma nota vinda do WhatsApp, pra devolver a
    confirmação (link + chave) pra quem pediu. Protegido por um segredo
    compartilhado (cabeçalho X-Responder-Secret) — sem ele, qualquer um
    que achasse essa URL conseguiria mandar mensagem usando o número do
    escritório."""
    if not WHATSAPP_PHONE_NUMBER_ID or not RESPONDER_SECRET:
        return jsonify({"erro": "WHATSAPP_PHONE_NUMBER_ID/RESPONDER_SECRET não configurados no Render."}), 503

    segredo_recebido = request.headers.get("X-Responder-Secret", "")
    if segredo_recebido != RESPONDER_SECRET:
        return jsonify({"erro": "segredo inválido"}), 403

    dados = request.get_json(silent=True) or {}
    telefone = dados.get("telefone", "")
    mensagem = dados.get("mensagem", "")
    if not telefone or not mensagem:
        return jsonify({"erro": "faltou 'telefone' e/ou 'mensagem' no corpo da requisição"}), 400

    try:
        resp = requests.post(
            f"{GRAPH_URL}/{WHATSAPP_PHONE_NUMBER_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"},
            json={
                "messaging_product": "whatsapp",
                "to": telefone,
                "type": "text",
                "text": {"body": mensagem},
            },
            timeout=20,
        )
        if resp.status_code >= 300:
            print(f"[ERRO] Meta recusou o envio pra {telefone}: {resp.status_code} {resp.text}")
            return jsonify({"erro": "Meta recusou o envio", "detalhe": resp.text}), 502
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] enviando WhatsApp pra {telefone}: {exc}")
        return jsonify({"erro": str(exc)}), 500

    return jsonify({"status": "enviado"}), 200


# =========================================================================
# PROCESSAMENTO DA MENSAGEM
# =========================================================================

def processar_mensagem(msg, contatos):
    remetente = msg.get("from", "")  # número do paciente/clínica, ex.: "5544999998888"
    nome_contato = contatos.get(remetente, "")
    tipo = msg.get("type")

    texto = ""
    anexos = []  # lista de (nome_arquivo, bytes, mime)

    if tipo == "text":
        texto = msg.get("text", {}).get("body", "")
    elif tipo in ("image", "document", "audio", "video"):
        bloco = msg.get(tipo, {})
        texto = bloco.get("caption", "")
        media_id = bloco.get("id")
        if media_id:
            resultado = baixar_midia(media_id)
            if resultado:
                anexos.append(resultado)
    else:
        texto = f"(tipo de mensagem não tratado pelo robô: {tipo})"

    enviar_email(remetente, nome_contato, texto, anexos)


def baixar_midia(media_id):
    """Duas chamadas, conforme a API da Meta exige: 1) pega a URL
    temporária + tipo do arquivo; 2) baixa o conteúdo de fato — as duas
    autenticadas com o mesmo token do app."""
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    try:
        resp = requests.get(f"{GRAPH_URL}/{media_id}", headers=headers, timeout=20)
        resp.raise_for_status()
        info = resp.json()
        url = info["url"]
        mime = info.get("mime_type", "application/octet-stream")

        resp2 = requests.get(url, headers=headers, timeout=30)
        resp2.raise_for_status()

        extensao = mime.split("/")[-1].split(";")[0]
        nome_arquivo = f"{media_id}.{extensao}"
        return nome_arquivo, resp2.content, mime
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] baixando mídia {media_id}: {exc}")
        return None


def enviar_email(remetente, nome_contato, texto, anexos):
    """Manda o e-mail de repasse via Microsoft Graph (/users/{...}/sendMail),
    com autenticação de aplicativo (client credentials) — SEM SMTP, porque
    o Render bloqueia conexões SMTP de saída."""
    quem = f"{nome_contato} ({remetente})" if nome_contato else remetente

    corpo_mensagem = {
        "subject": f"WhatsApp - {quem}",
        "body": {
            "contentType": "Text",
            "content": texto or "(mensagem sem texto — ver anexo)",
        },
        "toRecipients": [{"emailAddress": {"address": DESTINO_EMAIL}}],
    }

    if anexos:
        corpo_mensagem["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": nome_arquivo,
                "contentType": mime,
                "contentBytes": base64.b64encode(conteudo).decode("ascii"),
            }
            for nome_arquivo, conteudo, mime in anexos
        ]

    token = obter_token_graph()
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{REMETENTE_EMAIL}/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"message": corpo_mensagem, "saveToSentItems": "false"},
        timeout=30,
    )
    if resp.status_code >= 300:
        print(f"[ERRO] Graph recusou o envio do e-mail de repasse: {resp.status_code} {resp.text}")
        resp.raise_for_status()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
