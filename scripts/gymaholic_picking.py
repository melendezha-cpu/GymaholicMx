#!/usr/bin/env python3
"""
gymaholic_picking.py
Prototipo para job nocturno:
- Lee pedidos desde un JSON (samples/orders.json)
- Filtra las 4 familias (BULG, SANDBAG, FITBELL, BAZOTE + sets)
- Aplica excepciones y combos (SETFITBELL/SETBAZOTE, KITFUNC)
- Agrupa por marketplace (MercadoLibre primero, Shopify segundo)
- Genera texto de salida en el formato requerido
- Intenta enviar por SMTP (opcional) y/o por Telegram (requests o abrir navegador)
"""

import os
import json
import re
import logging
import argparse
from collections import defaultdict
from datetime import datetime
import smtplib
from email.message import EmailMessage
import time

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# --------------------
# Config / catálogo (simplificado; ampliar desde prompts.txt)
# --------------------
CATALOG_KEYWORDS = [
    ("balon medicinal medball crossfit 10 kg gymaholic", ("FITBELL10K","FITBELL",130)),
    ("fitbell balon medicinal", ("FITBELL","FITBELL",None)),
    ("balon medicinal medball", ("MEDBALL","FITBELL_OR_MEDBALL",None)),
    ("balon de azote", ("BAZOTE","BAZOTE",None)),
    ("set 4 balones azote", ("SETBAZOTE","SETBAZOTE",750)),
    ("set balones medicinales", ("SETFITBELL","SETFITBELL",550)),
    ("sandbag", ("SANDBAG","SANDBAG",None)),
    ("costal búlgaro", ("BULG","BULG",None)),
    ("kit funcional", ("KITFUNC","KITFUNC",None)),
]

COLOR_MAP = {"violeta":"MORADO","morado":"MORADO","rosa":"ROSA","negro":"NEGRO","rojo":"ROJO","azul":"AZUL","naranja":"NARANJA","amarillo":"AMARILLO"}

EMPLAYADO_CODES = {"SETFITBELL","SETBAZOTE"}

# --------------------
# Helpers
# --------------------
def normalize_color(raw):
    if not raw:
        return ""
    s = raw.strip().lower()
    s = re.sub(r'color[:\-]\s*','', s)
    s = s.split('|')[0].strip()
    return COLOR_MAP.get(s, s.upper())

def extract_weight_from_title(title):
    m = re.search(r'(\d{1,2})\s*kg', (title or "").lower())
    if m:
        return m.group(1)
    m2 = re.search(r'(\d{1,2})k\b', (title or "").lower())
    if m2:
        return m2.group(1)
    return ""

def detect_code_and_family(title, variant, publication_id=None):
    t = (title or "").strip()
    if t.lower() == "balon medicinal medball crossfit 10 kg gymaholic":
        return "FITBELL10K", "FITBELL"
    tl = t.lower()
    for kw, (code_base, family, cost) in CATALOG_KEYWORDS:
        if kw in tl:
            if code_base in ("MEDBALL","FITBELL"):
                if "agarradera" in tl or "agarrad" in (variant or "").lower():
                    weight = extract_weight_from_title(t) or ""
                    return f"FITBELL{weight}K" if weight else "FITBELL", "FITBELL"
                else:
                    weight = extract_weight_from_title(t) or ""
                    return f"MEDBALL{weight}K" if weight else "MEDBALL", "FITBELL_OR_MEDBALL"
            return code_base if code_base.endswith("K") or code_base.startswith("SET") or code_base=="KITFUNC" else code_base, family
    return None, None

def is_family_of_interest(code_or_family):
    if not code_or_family:
        return False
    for f in ["BULG","SANDBAG","FITBELL","BAZOTE","SETFITBELL","SETBAZOTE","KITFUNC"]:
        if f in code_or_family:
            return True
    return False

# --------------------
# Core processing
# --------------------
def load_orders(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def process_orders(orders):
    ml_agg = defaultdict(int)
    shopify_agg = defaultdict(int)
    revisar = []
    totals = {"ml":0,"shopify":0}
    for o in orders:
        src = (o.get("marketplace") or "").strip().lower()
        order_id = str(o.get("order_id") or "")
        items = o.get("items", [])
        for it in items:
            title = it.get("title") or it.get("titulo") or ""
            variant = it.get("variant") or it.get("variante") or ""
            qty = int(it.get("qty") or it.get("cantidad") or it.get("quantity") or 1)
            publication_id = it.get("publication_id") or it.get("publicacion")
            code, family = detect_code_and_family(title, variant, publication_id)
            color = normalize_color(variant)
            tl = (title or "").lower()
            if "set 4 balones azote" in tl or ("set" in tl and "azote" in tl):
                code = "SETBAZOTE"
                color = color or "SIN COLOR"
            if "set balones medicinales" in tl or "set fitbell" in tl:
                code = "SETFITBELL"
                color = color or "SIN COLOR"
            if "kit funcional" in tl or "kitfunc" in tl:
                # desglosar en 3 componentes
                comps = [("BULG20K", color or "SIN COLOR", 1), ("SANDBAG10K", color or "SIN COLOR", 1), ("FITBELL5K", color or "SIN COLOR", 1)]
                for ccode, ccolor, cq in comps:
                    key = (ccode, ccolor)
                    if src == "mercadolibre":
                        ml_agg[key] += qty * cq
                        totals["ml"] += qty * cq
                    else:
                        shopify_agg[key] += qty * cq
                        totals["shopify"] += qty * cq
                continue
            if not code or not is_family_of_interest(code):
                revisar.append({"order_id": order_id, "marketplace": src, "title": title, "variant": variant})
                continue
            if not code.endswith("K") and re.search(r'\d', code) is None:
                w = extract_weight_from_title(title)
                if w:
                    code = f"{code}{w}K"
            key = (code, color or "SIN COLOR")
            if src == "mercadolibre":
                ml_agg[key] += qty
                totals["ml"] += qty
            else:
                shopify_agg[key] += qty
                totals["shopify"] += qty
    return ml_agg, shopify_agg, revisar, totals

def build_text_output(ml_agg, shopify_agg, revisar, totals):
    lines = []
    lines.append("MERCADO LIBRE (prioridad - Por Enviar hoy + Próximos días):")
    total_ml = 0
    for (code,color),qty in sorted(ml_agg.items(), key=lambda x: x[0][0]):
        prod = f"{code} {color}".strip()
        if code in EMPLAYADO_CODES:
            prod = f"{code} {color} (EMPLAYADO)"
        lines.append(f"{qty} {prod}")
        total_ml += qty
    lines.append(f"TOTAL MERCADO LIBRE: {total_ml}")
    lines.append("")
    lines.append("SHOPIFY:")
    total_shop = 0
    for (code,color),qty in sorted(shopify_agg.items(), key=lambda x: x[0][0]):
        prod = f"{code} {color}".strip()
        if code in EMPLAYADO_CODES:
            prod = f"{code} {color} (EMPLAYADO)"
        lines.append(f"{qty} {prod}")
        total_shop += qty
    lines.append(f"TOTAL SHOPIFY: {total_shop}")
    lines.append("")
    lines.append(f"TOTAL GENERAL: {total_ml + total_shop}")
    lines.append("")
    lines.append("REVISAR MANUALMENTE:")
    if revisar:
        for r in revisar:
            lines.append(f"{r.get('marketplace','')} #{r.get('order_id')} — {r.get('title')} — {r.get('variant')}")
    else:
        lines.append("Ninguno")
    return "\n".join(lines)

# --------------------
# Sending utilities (SMTP + Telegram)
# --------------------
def send_email_smtp(subject, body, smtp_host, smtp_port, smtp_user, smtp_pass, from_addr, to_addr, use_tls=True, retries=3):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg.set_content(body)
    attempt = 0
    while attempt < retries:
        try:
            attempt += 1
            if use_tls:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as s:
                    s.starttls()
                    s.login(smtp_user, smtp_pass)
                    s.send_message(msg)
            else:
                with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) as s:
                    s.login(smtp_user, smtp_pass)
                    s.send_message(msg)
            logging.info("Email enviado correctamente a %s", to_addr)
            return True
        except Exception as e:
            logging.warning("Intento %s fallo al enviar correo: %s", attempt, e)
            time.sleep(2 ** attempt)
    return False

def send_telegram_via_requests(text, bot_token, chat_id):
    try:
        import requests
    except Exception:
        return False
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    r = requests.post(url, data={"chat_id": chat_id, "text": text})
    try:
        js = r.json()
        return js.get("ok", False)
    except Exception:
        return False

# --------------------
# CLI / main
# --------------------
def main():
    p = argparse.ArgumentParser()
    p.add_argument("--input", "-i", required=True, help="JSON file with orders")
    p.add_argument("--send-email", action="store_true", help="Enviar email via SMTP")
    p.add_argument("--send-telegram", action="store_true", help="Enviar Telegram (requests if possible)")
    args = p.parse_args()

    SMTP_HOST = os.getenv("SMTP_HOST")
    SMTP_PORT = int(os.getenv("SMTP_PORT","587"))
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")
    EMAIL_TO = os.getenv("EMAIL_TO","mlgymaholicmx@gmail.com")
    EMAIL_FROM = os.getenv("EMAIL_FROM", SMTP_USER or "gymaholicmx@gmail.com")
    TG_BOT = os.getenv("TELEGRAM_BOT_TOKEN")
    TG_CHAT = os.getenv("TELEGRAM_CHAT_ID")

    orders = load_orders(args.input)
    ml_agg, shopify_agg, revisar, totals = process_orders(orders)
    body = build_text_output(ml_agg, shopify_agg, revisar, totals)
    subj = f"Lista de picking GymaholicMx - {datetime.now().strftime('%d/%m/%Y')}"

    sent_email = False
    if args.send_email:
        if not all([SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, EMAIL_TO]):
            logging.error("Faltan variables SMTP: SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASS/EMAIL_TO")
        else:
            sent_email = send_email_smtp(subj, body, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, EMAIL_FROM, EMAIL_TO)

    sent_telegram = False
    if args.send_telegram:
        if TG_BOT and TG_CHAT:
            sent_telegram = send_telegram_via_requests(body, TG_BOT, TG_CHAT)
        else:
            logging.error("TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID no configurados")

    print("\n" + "="*60 + "\nOUTPUT (para chat / copia manual)\n" + "="*60 + "\n")
    print(body)
    print("\n" + "="*60 + "\nFIN\n" + "="*60 + "\n")

    logging.info("Resumen: EMAIL_SENT=%s TELEGRAM_SENT=%s ML_total=%s SHOP_total=%s revisar=%s",
                 sent_email, sent_telegram, sum(ml_agg.values()), sum(shopify_agg.values()), len(revisar))

if __name__ == "__main__":
    main()
