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


def _tax_de(it):
    """Extrae IVA (monto y %) de un ítem, sea tax lista, dict o nulo."""
    iva = 0.0
    iva_pct = 0.0
    tax = it.get("tax")
    if isinstance(tax, list):
        for t in tax:
            iva += _num(t.get("amount"))
            if _num(t.get("percentage")):
                iva_pct = _num(t.get("percentage"))
    elif isinstance(tax, dict):
        iva += _num(tax.get("amount"))
        iva_pct = _num(tax.get("percentage"))
    return iva, iva_pct


def bill_a_factura(bill):
    """Convierte una factura de proveedor de Alegra al formato del motor
    {nit, folio, prefijo, total, lineas:[{concepto, base, iva_pct, iva}]}."""
    prov = bill.get("provider") or {}
    ident = str(prov.get("identification") or "")
    nit = re.sub(r"\D", "", ident.split("-")[0])
    nt = bill.get("numberTemplate") or {}
    folio = str(nt.get("fullNumber") or nt.get("number") or bill.get("id") or "")

    # El detalle puede venir en purchases.categories, en items, o en categories
    purchases = bill.get("purchases") or {}
    fuente = purchases.get("categories") or bill.get("items") or bill.get("categories") or []

    lineas = []
    for it in fuente:
        base = _num(it.get("subtotal"))
        if not base:
            base = _num(it.get("total"))
        if not base:
            base = _num(it.get("price")) * _num(it.get("quantity") or 1)
        iva, iva_pct = _tax_de(it)
        lineas.append({"concepto": it.get("name") or it.get("description") or "",
                       "base": base, "iva_pct": iva_pct, "iva": iva})
    return {"nit": nit, "folio": folio, "prefijo": "", "total": _num(bill.get("total")),
            "lineas": lineas}


def get_bill(bill_id):
    """Trae el detalle completo de una factura de proveedor (incluye adjuntos con su URL si la hay)."""
    r = requests.get(BASE + "/bills/" + str(bill_id), headers=_headers(), timeout=40)
    r.raise_for_status()
    return r.json()


def diagnostico(desde=None, hasta=None):
    """Revisa si los XML adjuntos se pueden descargar por la API."""
    bills = get_bills(desde, hasta, max_bills=60)
    out = {"ok": True, "facturas_encontradas": len(bills)}
    if bills:
        out["ejemplo_convertido"] = bill_a_factura(bills[0])
        con_adj = next((b for b in bills if b.get("attachments")), None)
        if con_adj:
            try:
                detalle = get_bill(con_adj["id"])
                out["bill_con_adjunto_id"] = con_adj["id"]
                out["adjuntos_detalle"] = detalle.get("attachments")
            except Exception as e:
                out["adjuntos_error"] = str(e)
        else:
            out["adjuntos_detalle"] = "ninguna factura de la muestra tiene adjunto"
    return out


def facturas_desde_alegra(fecha_desde=None, fecha_hasta=None):
    """Devuelve {nit_folio: factura} igual que el lector de XML, para que el motor las use igual."""
    facturas = {}
    for b in get_bills(fecha_desde, fecha_hasta):
        f = bill_a_factura(b)
        if f["lineas"]:
            facturas[f["nit"] + "_" + f["folio"]] = f
            facturas[f["folio"]] = f
    return facturas
