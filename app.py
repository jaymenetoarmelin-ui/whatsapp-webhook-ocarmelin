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
  SMTP_USER         - ocarmelin@ocarmelin.com.br
  SMTP_PASSWORD     - senha (ou senha de app) da caixa ocarmelin@ocarmelin.com.br
  DESTINO_EMAIL     - opcional; se não definir, usa o próprio SMTP_USER
                      como destinatário.
"""

import os
import smtplib
from email.message import EmailMessage

import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

VERIFY_TOKEN = os.environ["VERIFY_TOKEN"]
WHATSAPP_TOKEN = os.environ["WHATSAPP_TOKEN"]
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]
DESTINO_EMAIL = os.environ.get("DESTINO_EMAIL", SMTP_USER)

GRAPH_URL = "https://graph.facebook.com/v25.0"


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
    msg = EmailMessage()
    quem = f"{nome_contato} ({remetente})" if nome_contato else remetente
    msg["Subject"] = f"WhatsApp - {quem}"
    msg["From"] = SMTP_USER
    msg["To"] = DESTINO_EMAIL
    msg.set_content(texto or "(mensagem sem texto — ver anexo)")

    for nome_arquivo, conteudo, mime in anexos:
        tipo_principal, subtipo = mime.split("/", 1)
        msg.add_attachment(conteudo, maintype=tipo_principal, subtype=subtipo, filename=nome_arquivo)

    with smtplib.SMTP("smtp.office365.com", 587) as servidor:
        servidor.starttls()
        servidor.login(SMTP_USER, SMTP_PASSWORD)
        servidor.send_message(msg)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
