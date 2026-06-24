"""Conexión con la API de Alegra: trae las facturas de proveedor (bills) con sus ítems.
Autenticación Basic (correo + token), guardados como variables de entorno en Render."""
import os
import re
import base64
import requests

BASE = "https://api.alegra.com/api/v1"


def _headers():
    email = os.environ.get("ALEGRA_EMAIL", "")
    token = os.environ.get("ALEGRA_TOKEN", "")
    raw = f"{email}:{token}".encode("utf-8")
    return {"Authorization": "Basic " + base64.b64encode(raw).decode("utf-8"),
            "Accept": "application/json"}


def hay_credenciales():
    return bool(os.environ.get("ALEGRA_EMAIL") and os.environ.get("ALEGRA_TOKEN"))


def get_bills(fecha_desde=None, fecha_hasta=None, max_bills=500):
    """Trae las facturas de proveedor de Alegra (paginado)."""
    headers = _headers()
    bills = []
    start = 0
    while len(bills) < max_bills:
        params = {"start": start, "limit": 30, "order_field": "date", "order_direction": "ASC"}
        if fecha_desde:
            params["date_afterwards"] = fecha_desde
        if fecha_hasta:
            params["date_before"] = fecha_hasta
        r = requests.get(BASE + "/bills", headers=headers, params=params, timeout=40)
        r.raise_for_status()
        page = r.json()
        if not isinstance(page, list) or not page:
            break
        bills.extend(page)
        if len(page) < 30:
            break
        start += 30
    return bills


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def bill_a_factura(bill):
    """Convierte una factura de proveedor de Alegra al formato del motor
    {nit, folio, prefijo, lineas:[{concepto, base, iva_pct, iva}]}."""
    prov = bill.get("provider") or {}
    ident = str(prov.get("identification") or "")
    nit = re.sub(r"\D", "", ident.split("-")[0])
    nt = bill.get("numberTemplate") or {}
    folio = str(nt.get("fullNumber") or nt.get("number") or bill.get("id") or "")
    lineas = []
    for it in bill.get("items", []) or []:
        base = _num(it.get("total"))
        if not base:
            base = _num(it.get("price")) * _num(it.get("quantity") or 1)
        iva = 0.0
        iva_pct = 0.0
        for t in it.get("tax", []) or []:
            iva += _num(t.get("amount"))
            if _num(t.get("percentage")):
                iva_pct = _num(t.get("percentage"))
        lineas.append({"concepto": it.get("name") or it.get("description") or "",
                       "base": base, "iva_pct": iva_pct, "iva": iva})
    return {"nit": nit, "folio": folio, "prefijo": "", "lineas": lineas}


def facturas_desde_alegra(fecha_desde=None, fecha_hasta=None):
    """Devuelve {nit_folio: factura} igual que el lector de XML, para que el motor las use igual."""
    facturas = {}
    for b in get_bills(fecha_desde, fecha_hasta):
        f = bill_a_factura(b)
        if f["lineas"]:
            facturas[f["nit"] + "_" + f["folio"]] = f
            facturas[f["folio"]] = f
    return facturas
