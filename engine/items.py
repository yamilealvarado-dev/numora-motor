"""Clasificación ítem por ítem: a cada línea del XML/Alegra le sugiere su cuenta y su cuenta de IVA.
El usuario revisa y edita en la pantalla; aquí solo se sugiere."""
from collections import defaultdict
from engine.reglas import cuenta_por_concepto


def _limpiar(c):
    return str(c).strip() if c else ''


def sugerir_cuenta_item(linea, perfil, reglas=None):
    """Sugiere la cuenta de un ítem: 1) palabra clave del concepto, 2) cuenta del proveedor."""
    # 1) por concepto (separa aseo, cafetería, etc. dentro de una factura mixta)
    c = cuenta_por_concepto(linea.get('concepto', ''), reglas)
    if c:
        return _limpiar(c), 'concepto'
    # 2) por proveedor (la cuenta que más usa ese proveedor)
    if perfil:
        if linea.get('iva', 0) > 0 and perfil.get('grav_base'):
            return _limpiar(perfil['grav_base']), 'proveedor'
        if perfil.get('nograv'):
            return _limpiar(perfil['nograv']), 'proveedor'
        if perfil.get('grav_base'):
            return _limpiar(perfil['grav_base']), 'proveedor'
    return '', 'sin_sugerencia'


def clasificar_factura(factura, perfil, iva_pares, reglas=None, cuenta_iva_defecto='240820'):
    """Devuelve la factura con cada ítem clasificado (cuenta + cuenta de IVA sugeridas)
    y el asiento contable armado (partida doble)."""
    items = []
    for ln in factura.get('lineas', []):
        cuenta, origen = sugerir_cuenta_item(ln, perfil, reglas)
        iva = ln.get('iva', 0) or 0
        # cuenta de IVA: la pareja aprendida de esa cuenta; si no hay, descontable por defecto
        iva_cuenta = iva_pares.get(cuenta, cuenta_iva_defecto) if iva else ''
        items.append({
            'concepto': ln.get('concepto', ''),
            'base': ln.get('base', 0) or 0,
            'iva': iva,
            'iva_pct': ln.get('iva_pct', 0),
            'cuenta': cuenta,
            'iva_cuenta': iva_cuenta,
            'origen_sugerencia': origen,
            'auto': bool(cuenta),
        })
    asiento = armar_asiento(items, perfil)
    return {
        'nit': factura.get('nit', ''),
        'folio': factura.get('folio', ''),
        'prefijo': factura.get('prefijo', ''),
        'nombre': (perfil or {}).get('nombre', ''),
        'cxp': _limpiar((perfil or {}).get('prov', '220501')),
        'items': items,
        'asiento': asiento,
        'cuadra': asiento['cuadra'],
        'sin_cuenta': sum(1 for it in items if not it['cuenta']),
    }


def armar_asiento(items, perfil, cxp=None):
    """Agrupa los ítems en líneas contables: Db por cuenta (base) + Db por cuenta IVA + Cr proveedor."""
    debitos = defaultdict(float)
    for it in items:
        if it['cuenta']:
            debitos[it['cuenta']] += it['base']
        if it['iva']:
            k = it['iva_cuenta'] or it['cuenta']
            debitos[k] += it['iva']
    total_deb = sum(debitos.values())
    cxp = _limpiar(cxp or (perfil or {}).get('prov', '220501'))
    lineas = [{'cuenta': c, 'debito': round(v, 2), 'credito': 0} for c, v in debitos.items()]
    lineas.append({'cuenta': cxp, 'debito': 0, 'credito': round(total_deb, 2)})
    total_cred = total_deb
    return {'lineas': lineas, 'total_debito': round(total_deb, 2),
            'total_credito': round(total_cred, 2),
            'cuadra': abs(total_deb - total_cred) < 1}
