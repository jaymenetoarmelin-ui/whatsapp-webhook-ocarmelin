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

NOVO (12/11/2026) — FILA NO POSTGRES: cada pedido também é gravado numa
fila (banco Postgres do Render) e o robô do servidor busca direto dela
(GET /fila), sem depender do e-mail da Microsoft. O e-mail continua sendo
mandado como reserva; o robô descarta a cópia repetida. Variáveis:

  DATABASE_URL      - a "Internal Database URL" do Postgres do Render.
                      Sem ela, tudo funciona como antes (só e-mail).
  ALERTA_TELEFONE   - WhatsApp que recebe o aviso quando o robô fica
                      mais de ALERTA_MINUTOS (padrão 15) sem buscar a fila
                      (ex.: servidor desligou). Padrão: 554498357733.

O robô se identifica com o mesmo RESPONDER_SECRET do /responder.
"""

import base64
import json
import os
import re
import threading
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

# NOVO (02/10/2026, a pedido) — evita processar a MESMA mensagem do
# WhatsApp duas vezes. Causa real observada em produção: quando o webhook
# demora a responder (ou dá erro) — como aconteceu no apagão do SMTP —, a
# Meta reenvia a mensagem original automaticamente mais tarde (minutos a
# horas depois), com o mesmo "id" de mensagem (wamid). Sem essa checagem,
# o reenvio virava um e-mail de repasse duplicado, capturado pelo monitor
# como se fosse nota nova — e quase gerou uma SEGUNDA NFS-e da mesma nota.
# Cache simples em memória (zera se o serviço reiniciar — risco residual
# aceitável: a Meta reenvia pouquíssimas vezes e o serviço, agora num
# plano pago, não fica mais "dormindo"/reiniciando por inatividade).
_mensagens_vistas = set()
_mensagens_vistas_ordem = []
_LIMITE_CACHE_DEDUP = 1000
_lock_dedup = threading.Lock()


def ja_processada(msg_id):
    """True se esse wamid já foi visto antes (reenvio da Meta) — nesse
    caso o chamador deve pular a mensagem, sem mandar e-mail de novo."""
    if not msg_id:
        return False
    with _lock_dedup:
        if msg_id in _mensagens_vistas:
            return True
        _mensagens_vistas.add(msg_id)
        _mensagens_vistas_ordem.append(msg_id)
        if len(_mensagens_vistas_ordem) > _LIMITE_CACHE_DEDUP:
            mais_antigo = _mensagens_vistas_ordem.pop(0)
            _mensagens_vistas.discard(mais_antigo)
        return False

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
                    if ja_processada(msg.get("id")):
                        print(f"[INFO] Mensagem {msg.get('id')} já processada antes "
                              f"(reenvio da Meta) — ignorando pra não duplicar.")
                        continue
                    processar_mensagem(msg, contatos)
    except Exception as exc:  # noqa: BLE001 — nunca pode derrubar o endpoint
        print(f"[ERRO] processando webhook: {exc}")
    return jsonify({"status": "ok"}), 200


@app.route("/")
def raiz():
    return "OK — webhook do robô de captura WhatsApp está no ar."


def enviar_whatsapp(telefone, mensagem):
    """Manda texto pelo número do escritório. Devolve (ok, detalhe)."""
    if not WHATSAPP_PHONE_NUMBER_ID:
        return False, "WHATSAPP_PHONE_NUMBER_ID não configurado"
    try:
        resp = requests.post(
            f"{GRAPH_URL}/{WHATSAPP_PHONE_NUMBER_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"},
            json={"messaging_product": "whatsapp", "to": telefone, "type": "text",
                  "text": {"body": mensagem}},
            timeout=20,
        )
        if resp.status_code >= 300:
            print(f"[ERRO] Meta recusou o envio pra {telefone}: {resp.status_code} {resp.text}")
            return False, resp.text
        return True, ""
    except Exception as exc:  # noqa: BLE001
        print(f"[ERRO] enviando WhatsApp pra {telefone}: {exc}")
        return False, str(exc)


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

# Palavra-chave que marca um pedido de nota (a mesma regra do robô no PC:
# tolera maiúscula/minúscula, "NFSe" sem hífen e espaços a mais).
RE_PALAVRA_CHAVE = re.compile(r"Emitir\s+NFS\s*-?\s*e\b", re.IGNORECASE)

# NOVO (16/10/2026) — comandos de CANCELAMENTO de NFS-e (tratados pelo
# cancelar_nfse.py no servidor). "CANCELAR ..." é repassado de qualquer
# número (o robô recusa e registra quem não é autorizado); as respostas
# curtas "SIM <nº>" / "NÃO" só são repassadas dos números autorizados,
# pra conversa comum ("sim", "não") não encher a caixa de e-mail.
RE_CANCELAR = re.compile(r"^[ \t*]*CANCELAR\b", re.IGNORECASE | re.MULTILINE)
RE_CONFIRMACAO = re.compile(r"^\s*\**\s*(SIM(\s+\d+)?|N[ÃA]O)\s*[.!*]*\s*$", re.IGNORECASE)
TELEFONES_CANCELAMENTO = [
    re.sub(r"\D", "", t) for t in
    os.environ.get("CANCELAMENTO_TELEFONES_AUTORIZADOS", "554498357733").split(",")
    if re.sub(r"\D", "", t)
]


def eh_comando_cancelamento(texto, remetente):
    if RE_CANCELAR.search(texto or ""):
        return True
    tel = re.sub(r"\D", "", remetente or "")
    autorizado = any(tel.endswith(a) or a.endswith(tel) for a in TELEFONES_CANCELAMENTO) if tel else False
    # CORRIGIDO (12/11/2026): aceita várias confirmações numa mensagem só
    # ("SIM 29" e "SIM 30", uma por linha) — antes só passava uma.
    linhas = [l for l in (texto or "").splitlines() if l.strip()]
    return autorizado and bool(linhas) and all(RE_CONFIRMACAO.match(l) for l in linhas)


def processar_mensagem(msg, contatos):
    remetente = msg.get("from", "")  # número do paciente/clínica, ex.: "5544999998888"
    nome_contato = contatos.get(remetente, "")
    tipo = msg.get("type")

    texto = ""
    anexos = []  # lista de (nome_arquivo, bytes, mime)
    media_id = None

    if tipo == "text":
        texto = msg.get("text", {}).get("body", "")
    elif tipo in ("image", "document", "audio", "video"):
        bloco = msg.get(tipo, {})
        texto = bloco.get("caption", "")
        media_id = bloco.get("id")

    # NOVO (05/10/2026, a pedido) — só vira e-mail o que for PEDIDO DE
    # NOTA: mensagem cujo texto (ou legenda do arquivo) contenha a
    # palavra-chave "Emitir NFS-e". Conversa comum, arquivo sem legenda,
    # áudio, figurinha etc. são descartados aqui mesmo, sem e-mail — antes
    # TUDO que chegava no número era repassado e a caixa enchia de
    # conversa aleatória. O arquivo só é baixado da Meta depois dessa
    # checagem.
    if not RE_PALAVRA_CHAVE.search(texto or "") and not eh_comando_cancelamento(texto, remetente):
        print(f"[INFO] Mensagem de {remetente} (tipo {tipo}) sem 'Emitir NFS-e' nem "
              f"comando de cancelamento — descartada, sem e-mail.")
        return

    # NOVO (12/11/2026): grava na fila ANTES de tudo. O e-mail vira reserva:
    # se a Microsoft cair, o robô ainda pega o pedido pela fila.
    bloco = msg.get(tipo, {}) if media_id else {}
    fila_gravar(msg.get("id"), remetente, nome_contato, texto, media_id,
                bloco.get("mime_type", ""), bloco.get("filename", ""))

    if media_id:
        resultado = baixar_midia(media_id)
        if resultado:
            anexos.append(resultado)

    try:
        enviar_email(remetente, nome_contato, texto, anexos)
    except Exception as exc:  # noqa: BLE001 — com a fila, e-mail é só reserva
        print(f"[ERRO] e-mail de repasse falhou (o pedido está na fila): {exc}")


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


# =========================================================================
# NOVO (12/11/2026) — FILA NO POSTGRES
# =========================================================================

DATABASE_URL = os.environ.get("DATABASE_URL", "")
ALERTA_TELEFONE = re.sub(r"\D", "", os.environ.get("ALERTA_TELEFONE", "554498357733"))
ALERTA_MINUTOS = int(os.environ.get("ALERTA_MINUTOS", "15"))
DIAS_GUARDAR = int(os.environ.get("FILA_DIAS_GUARDAR", "90"))


def _conexao():
    import psycopg  # só importado quando há banco configurado
    return psycopg.connect(DATABASE_URL, autocommit=True, connect_timeout=10)


def fila_preparar():
    if not DATABASE_URL:
        print("[INFO] DATABASE_URL não configurada — fila desligada (só e-mail).")
        return
    with _conexao() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS fila_whatsapp (
                wamid         TEXT PRIMARY KEY,
                recebido_em   TIMESTAMPTZ NOT NULL DEFAULT now(),
                remetente     TEXT NOT NULL,
                nome          TEXT,
                texto         TEXT,
                media_id      TEXT,
                media_mime    TEXT,
                media_nome    TEXT,
                confirmado_em TIMESTAMPTZ
            )""")
        c.execute("CREATE INDEX IF NOT EXISTS fila_pendentes ON fila_whatsapp (recebido_em) "
                  "WHERE confirmado_em IS NULL")
        c.execute("""
            CREATE TABLE IF NOT EXISTS robo_status (
                id              INT PRIMARY KEY,
                ultimo_contato  TIMESTAMPTZ,
                alertado        BOOLEAN NOT NULL DEFAULT FALSE
            )""")
        # ultimo_contato começa vazio: o vigia só passa a cobrar o robô depois
        # da primeira busca dele na fila (sem alarme falso antes disso).
        c.execute("INSERT INTO robo_status (id, ultimo_contato) VALUES (1, NULL) ON CONFLICT DO NOTHING")
    print("[INFO] Fila no Postgres pronta.")


def fila_gravar(wamid, remetente, nome, texto, media_id, media_mime, media_nome):
    if not DATABASE_URL or not wamid:
        return
    try:
        with _conexao() as c:
            c.execute("INSERT INTO fila_whatsapp (wamid, remetente, nome, texto, media_id, media_mime, "
                      "media_nome) VALUES (%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (wamid) DO NOTHING",
                      (wamid, remetente, nome, texto, media_id, media_mime, media_nome))
    except Exception as exc:  # noqa: BLE001 — o e-mail ainda leva o pedido
        print(f"[ERRO] gravando {wamid} na fila: {exc}")


def _robo_autenticado():
    return bool(RESPONDER_SECRET) and request.headers.get("X-Responder-Secret", "") == RESPONDER_SECRET


@app.route("/fila", methods=["GET"])
def fila_listar():
    """O robô busca aqui, a cada ciclo, os pedidos ainda não confirmados.
    Cada chamada também conta como 'sinal de vida' do robô."""
    if not _robo_autenticado():
        return jsonify({"erro": "segredo inválido"}), 403
    if not DATABASE_URL:
        return jsonify({"erro": "fila desligada (DATABASE_URL)"}), 503
    with _conexao() as c:
        c.execute("UPDATE robo_status SET ultimo_contato = now() WHERE id = 1")
        cur = c.execute("SELECT wamid, recebido_em, remetente, nome, texto, media_id, media_mime, media_nome "
                        "FROM fila_whatsapp WHERE confirmado_em IS NULL ORDER BY recebido_em LIMIT 50")
        itens = [{"wamid": r[0], "recebido_em": r[1].isoformat(), "remetente": r[2], "nome": r[3] or "",
                  "texto": r[4] or "", "media_id": r[5], "media_mime": r[6], "media_nome": r[7]}
                 for r in cur.fetchall()]
    return jsonify({"itens": itens}), 200


@app.route("/fila/confirmar", methods=["POST"])
def fila_confirmar():
    """O robô avisa quais pedidos já tratou (gravou na planilha ou descartou)."""
    if not _robo_autenticado():
        return jsonify({"erro": "segredo inválido"}), 403
    ids = (request.get_json(silent=True) or {}).get("wamids") or []
    if ids:
        with _conexao() as c:
            c.execute("UPDATE fila_whatsapp SET confirmado_em = now() "
                      "WHERE wamid = ANY(%s) AND confirmado_em IS NULL", (list(ids),))
    return jsonify({"status": "ok", "confirmados": len(ids)}), 200


@app.route("/fila/midia/<media_id>", methods=["GET"])
def fila_midia(media_id):
    """Entrega ao robô o anexo (ex.: planilha de Agendamentos) baixando
    da Meta na hora — o arquivo não fica guardado no banco."""
    if not _robo_autenticado():
        return jsonify({"erro": "segredo inválido"}), 403
    resultado = baixar_midia(media_id)
    if not resultado:
        return jsonify({"erro": "não consegui baixar o arquivo na Meta"}), 502
    nome, conteudo, mime = resultado
    return conteudo, 200, {"Content-Type": mime, "X-Nome-Arquivo": nome}


def _vigia():
    """Roda em segundo plano: avisa no WhatsApp se o robô parar de buscar
    a fila, avisa quando ele volta, e limpa pedidos antigos já tratados.
    O UPDATE ... RETURNING garante um aviso só, mesmo com vários workers."""
    ultima_limpeza = 0
    while True:
        time.sleep(60)
        try:
            with _conexao() as c:
                cur = c.execute(
                    "UPDATE robo_status SET alertado = TRUE WHERE id = 1 AND NOT alertado "
                    "AND ultimo_contato < now() - make_interval(mins => %s::int) RETURNING ultimo_contato",
                    (ALERTA_MINUTOS,))
                parou = cur.fetchone()
                if parou:
                    pendentes = c.execute("SELECT count(*) FROM fila_whatsapp "
                                          "WHERE confirmado_em IS NULL").fetchone()[0]
                    enviar_whatsapp(ALERTA_TELEFONE, (
                        f"⚠️ Robô NFS-e PARADO: o servidor não busca pedidos há mais de {ALERTA_MINUTOS} min. "
                        f"{pendentes} pedido(s) aguardando. Confira se o srv-armelin está ligado e rode "
                        "Ativar_Robo.bat — os pedidos ficam guardados e são emitidos quando ele voltar."))
                cur = c.execute(
                    "UPDATE robo_status SET alertado = FALSE WHERE id = 1 AND alertado "
                    "AND ultimo_contato > now() - interval '2 minutes' RETURNING id")
                if cur.fetchone():
                    enviar_whatsapp(ALERTA_TELEFONE, "✅ Robô NFS-e voltou a buscar os pedidos.")
                if time.time() - ultima_limpeza > 6 * 3600:
                    c.execute("DELETE FROM fila_whatsapp WHERE confirmado_em < now() - make_interval(days => %s::int)",
                              (DIAS_GUARDAR,))
                    ultima_limpeza = time.time()
        except Exception as exc:  # noqa: BLE001 — o vigia nunca pode morrer
            print(f"[ERRO] vigia da fila: {exc}")


if DATABASE_URL:
    try:
        fila_preparar()
        threading.Thread(target=_vigia, daemon=True).start()
    except Exception as exc:  # noqa: BLE001 — sem banco, o webhook segue com e-mail
        print(f"[ERRO] preparando a fila no Postgres: {exc}")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
