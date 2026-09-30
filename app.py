#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Webhook do WhatsApp Business Cloud API (Meta) -> repassa cada mensagem
recebida (texto + anexos: foto, PDF, áudio) como um e-mail pra
ocarmelin@ocarmelin.com.br, via Microsoft Graph API (não SMTP — o Render
bloqueia a porta SMTP de saída), reaproveitando o mesmo fluxo de captura
de dados que já lê e-mails de outros remetentes.

Variáveis de ambiente necessárias no Render:
  VERIFY_TOKEN     - string qualquer, usada na verificação do webhook na Meta.
  WHATSAPP_TOKEN   - token PERMANENTE do app da Meta (Usuário do sistema).
  MS_TENANT_ID     - ID do diretório (tenant) do Azure AD.
  MS_CLIENT_ID     - ID do aplicativo (client id) registrado no Azure AD.
  MS_CLIENT_SECRET - segredo do cliente gerado no Azure AD.
  REMETENTE_EMAIL  - ocarmelin@ocarmelin.com.br (caixa que manda e recebe).
  DESTINO_EMAIL    - opcional; se não definir, usa REMETENTE_EMAIL.
"""

import base64
import os

import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

VERIFY_TOKEN = os.environ["VERIFY_TOKEN"]
WHATSAPP_TOKEN = os.environ["WHATSAPP_TOKEN"]
MS_TENANT_ID = os.environ["MS_TENANT_ID"]
MS_CLIENT_ID = os.environ["MS_CLIENT_ID"]
MS_CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
REMETENTE_EMAIL = os.environ["REMETENTE_EMAIL"]
DESTINO_EMAIL = os.environ.get("DESTINO_EMAIL", REMETENTE_EMAIL)

GRAPH_URL_WHATSAPP = "https://graph.facebook.com/v25.0"


@app.route("/webhook", methods=["GET"])
def verificar_webhook():
    modo = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if modo == "subscribe" and token == VERIFY_TOKEN:
        return challenge, 200
    return "Verificação falhou — token não confere.", 403


@app.route("/webhook", methods=["POST"])
def receber_webhook():
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
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] processando webhook: {exc}")
    return jsonify({"status": "ok"}), 200


@app.route("/privacidade")
def politica_privacidade():
    return """
    <h1>Política de Privacidade — Captura Dados (OCArmelin)</h1>
    <p>Este aplicativo é de uso interno da Organização Contábil Armelin e
    recebe mensagens do WhatsApp enviadas por clientes (pacientes/clínicas)
    para fins de emissão de nota fiscal de serviço (NFS-e).</p>
    <p>As mensagens recebidas (texto, imagens e documentos) são
    encaminhadas por e-mail para a caixa interna da contabilidade
    (ocarmelin@ocarmelin.com.br) e usadas apenas para identificar dados
    da nota fiscal a ser emitida (nome, valor, data do serviço).</p>
    <p>Não compartilhamos esses dados com terceiros, exceto quando
    exigido pela emissão da própria nota fiscal junto à Receita Federal /
    prefeituras (Sistema Nacional NFS-e).</p>
    <p>Para dúvidas ou solicitação de exclusão de dados, entre em
    contato: ocarmelin@ocarmelin.com.br</p>
    """


@app.route("/")
def raiz():
    return "OK — webhook do robô de captura WhatsApp está no ar."


# =========================================================================
# PROCESSAMENTO DA MENSAGEM
# =========================================================================

def processar_mensagem(msg, contatos):
    remetente = msg.get("from", "")
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
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    try:
        resp = requests.get(f"{GRAPH_URL_WHATSAPP}/{media_id}", headers=headers, timeout=20)
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


# =========================================================================
# ENVIO DE E-MAIL VIA MICROSOFT GRAPH (não usa SMTP — porta bloqueada no Render)
# =========================================================================

_cache_token = {"valor": None}


def obter_token_graph():
    """Client Credentials flow — token de aplicativo, sem usuário logado."""
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    dados = {
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=dados, timeout=20)
    resp.raise_for_status()
    return resp.json()["access_token"]


def enviar_email(remetente, nome_contato, texto, anexos):
    token = obter_token_graph()
    quem = f"{nome_contato} ({remetente})" if nome_contato else remetente

    anexos_graph = []
    for nome_arquivo, conteudo, mime in anexos:
        anexos_graph.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": nome_arquivo,
            "contentType": mime,
            "contentBytes": base64.b64encode(conteudo).decode("ascii"),
        })

    corpo = {
        "message": {
            "subject": f"WhatsApp - {quem}",
            "body": {"contentType": "Text", "content": texto or "(mensagem sem texto — ver anexo)"},
            "toRecipients": [{"emailAddress": {"address": DESTINO_EMAIL}}],
            "attachments": anexos_graph,
        },
        "saveToSentItems": "true",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{REMETENTE_EMAIL}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=corpo, timeout=30)
    if resp.status_code >= 300:
        print(f"[ERRO] Graph sendMail falhou ({resp.status_code}): {resp.text[:500]}")
    resp.raise_for_status()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
