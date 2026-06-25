"""Lee los XML de facturas electrónicas (formato DIAN/UBL) y extrae las líneas (concepto, base, IVA)."""
import os
import re
import glob
import zipfile
import tempfile
from lxml import etree

NS = {'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
      'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'}


def _invoice_root(path):
    try:
        root = etree.fromstring(open(path, 'rb').read())
    except Exception:
        return None
    tag_desc = '{urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2}Description'
    for desc in root.iter(tag_desc):
        if desc.text and ('<Invoice' in desc.text or '<CreditNote' in desc.text or 'cbc:ID' in desc.text):
            try:
                return etree.fromstring(desc.text.encode('utf-8'))
            except Exception:
                pass
    if 'Invoice' in root.tag or 'CreditNote' in root.tag:
        return root
    return None


def _f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def leer_factura(path):
    """Devuelve dict {nit, folio, prefijo, lineas:[{concepto, base, iva_pct, iva}]} o None."""
    inv = _invoice_root(path)
    if inv is None:
        return None
    emisor_nit = inv.find('.//cac:AccountingSupplierParty//cbc:CompanyID', NS)
    folio = inv.find('cbc:ID', NS)
    fecha = inv.find('cbc:IssueDate', NS)
    fecha_txt = ''
    if fecha is not None and fecha.text:
        try:
            y, m, d = fecha.text.strip()[:10].split('-')
            fecha_txt = f"{d}/{m}/{y}"
        except ValueError:
            fecha_txt = fecha.text.strip()
    pref = ''
    fid = folio.text if folio is not None else ''
    m = re.match(r'([A-Za-z]+)(\d+)', fid or '')
    if m:
        pref, fid = m.group(1), m.group(2)
    lineas = []
    line_tag = 'cac:InvoiceLine' if inv.findall('.//cac:InvoiceLine', NS) else 'cac:CreditNoteLine'
    for ln in inv.findall('.//' + line_tag, NS):
        d = ln.find('.//cbc:Description', NS)
        amt = ln.find('cbc:LineExtensionAmount', NS)
        tax = ln.find('.//cac:TaxTotal/cbc:TaxAmount', NS)
        pct = ln.find('.//cac:TaxSubtotal//cbc:Percent', NS)
        lineas.append({'concepto': (d.text or '').strip() if d is not None else '',
                       'base': _f(amt.text if amt is not None else 0),
                       'iva_pct': _f(pct.text if pct is not None else 0),
                       'iva': _f(tax.text if tax is not None else 0)})
    return {'nit': re.sub(r'\D', '', emisor_nit.text) if emisor_nit is not None else '',
            'folio': fid, 'prefijo': pref, 'fecha': fecha_txt, 'lineas': lineas}


def leer_zip(ruta_zip):
    """Descomprime un ZIP (con ZIPs anidados) y devuelve {nit_folio: factura}."""
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(ruta_zip) as z:
        z.extractall(tmp)
    for zp in glob.glob(os.path.join(tmp, '**', '*.zip'), recursive=True):
        try:
            with zipfile.ZipFile(zp) as z2:
                z2.extractall(os.path.join(tmp, 'nested'))
        except Exception:
            pass
    facturas = {}
    for xml in glob.glob(os.path.join(tmp, '**', '*.xml'), recursive=True):
        f = leer_factura(xml)
        if f and f['lineas']:
            facturas[f['nit'] + '_' + f['folio']] = f
            facturas[f['folio']] = f
    return facturas
