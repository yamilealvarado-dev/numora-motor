"""Toma el reporte DIAN + perfiles + (opcional) líneas de XML y arma el asiento de cada compra."""
import pandas as pd
from .aprender import limpiar_nit
from .reglas import cuenta_por_concepto
from .items import clasificar_factura, _limpiar


def _num(x):
    return pd.to_numeric(x, errors='coerce')


def clasificar_dian_items(ruta_dian, perfiles, iva_pares, facturas_xml=None, reglas=None):
    """Versión ítem por ítem: el reporte DIAN es la lista maestra; el XML da el detalle.
    Cada asiento trae 'items' (detalle editable) y 'lineas' (el asiento contable)."""
    facturas_xml = facturas_xml or {}
    iva_pares = iva_pares or {}
    df = pd.read_excel(ruta_dian, header=0, dtype=str)
    rec = df[df['Grupo'].astype(str).str.contains('ecib', na=False)].copy()
    for c in ['IVA', 'INC', 'Total']:
        rec[c] = _num(rec[c]).fillna(0)
    rec['nit'] = rec['NIT Emisor'].apply(limpiar_nit)
    rec['fe'] = pd.to_datetime(rec['Fecha Emisión'], format='%d-%m-%Y', errors='coerce')
    rec = rec.sort_values('fe')

    asientos = []
    for _, r in rec.iterrows():
        nota = 'rédito' in str(r['Tipo de documento'])
        signo = -1 if nota else 1
        T, V = r['Total'], r['IVA']
        nit = r['nit']
        p = perfiles.get(nit)
        folio = f"{r['Prefijo'] or ''}{r['Folio'] or ''}"
        clave = nit + '_' + str(r['Folio'] or '')
        fac = facturas_xml.get(clave) or facturas_xml.get(str(r['Folio'] or ''))
        fecha = r['fe'].strftime('%d/%m/%Y') if pd.notna(r['fe']) else ''
        nombre = str(r['Nombre Emisor'])[:40]

        if fac and fac.get('lineas'):
            # CON XML: detalle ítem por ítem
            rr = clasificar_factura(fac, p, iva_pares, reglas)
            items = rr['items']
            asi_lineas = rr['asiento']['lineas']
            prov_cuenta = _limpiar(rr['cxp'])
            origen = 'XML'
            estado = 'OK' if rr['sin_cuenta'] == 0 else f"{rr['sin_cuenta']} ítem(s) sin cuenta"
        else:
            # SIN XML: un solo ítem a nivel factura (la base + su IVA)
            base = T - V
            if p:
                cta = _limpiar(p['grav_base'] or p['nograv']) if V > 0 else _limpiar(p['nograv'] or p['grav_base'])
                iva_cta = iva_pares.get(cta, p.get('iva') or '240820') if V else ''
                prov_cuenta = _limpiar(p['prov'])
                estado = 'Revisar (sin XML)' if len(p['cuentas']) > 1 else 'OK'
            else:
                cta = ''
                iva_cta = '240820' if V else ''
                prov_cuenta = '220501'
                estado = 'NUEVO - asignar cuenta'
            items = [{'concepto': f'(sin XML) {nombre}', 'base': round(base, 2), 'iva': round(V, 2),
                      'iva_pct': 0, 'cuenta': _limpiar(cta), 'iva_cuenta': _limpiar(iva_cta),
                      'origen_sugerencia': 'proveedor' if p else 'sin_sugerencia', 'auto': bool(cta)}]
            from .items import armar_asiento
            asi = armar_asiento(items, p, prov_cuenta)
            asi_lineas = asi['lineas']
            origen = 'sin XML'

        # separar débitos (van en 'lineas') del crédito al proveedor
        lineas = [{'cuenta': l['cuenta'], 'debito': round(l['debito'] * signo, 2)}
                  for l in asi_lineas if l['debito']]
        credito = round(sum(l['credito'] for l in asi_lineas) * signo, 2)
        total_deb = round(sum(l['debito'] for l in lineas), 2)

        asientos.append({
            'fecha': fecha, 'tipo': 'NC' if nota else 'FE',
            'factura': folio, 'nit': nit, 'proveedor': nombre,
            'total': round(T * signo, 2), 'iva': round(V * signo, 2),
            'origen': origen, 'items': items, 'lineas': lineas,
            'prov_cuenta': prov_cuenta, 'credito': credito, 'estado': estado,
            'comprobante': '00003', 'dividida': len(lineas) > 2,
            'cuadra': abs(total_deb - credito) < 1,
            'sin_cuenta': sum(1 for it in items if not it['cuenta']),
        })

    resumen = {
        'facturas': len(asientos),
        'total': round(sum(a['total'] for a in asientos), 2),
        'ok': sum(1 for a in asientos if a['estado'] == 'OK'),
        'con_xml': sum(1 for a in asientos if a['origen'] == 'XML'),
        'divididas': sum(1 for a in asientos if a['origen'] == 'XML' and len(a['lineas']) > 2),
        'nuevos': sum(1 for a in asientos if 'NUEVO' in a['estado']),
        'revisar': sum(1 for a in asientos if 'Revisar' in a['estado'] or 'sin cuenta' in a['estado']),
        'descuadres': sum(1 for a in asientos if not a['cuadra']),
        'items_total': sum(len(a['items']) for a in asientos),
    }
    return asientos, resumen


def clasificar_solo_xml(facturas_xml, perfiles, iva_pares, reglas=None):
    """Modo SOLO XML: sin reporte DIAN, los XML son la lista. Cada uno se clasifica ítem por ítem."""
    iva_pares = iva_pares or {}
    # dedup: leer_zip guarda dos llaves por factura
    vistos = set()
    unicas = []
    for k, f in (facturas_xml or {}).items():
        key = f['nit'] + '_' + f['folio']
        if key not in vistos:
            vistos.add(key)
            unicas.append(f)

    asientos = []
    for f in unicas:
        p = perfiles.get(f['nit'])
        rr = clasificar_factura(f, p, iva_pares, reglas)
        asi_lineas = rr['asiento']['lineas']
        lineas = [{'cuenta': l['cuenta'], 'debito': l['debito']} for l in asi_lineas if l['debito']]
        credito = round(sum(l['credito'] for l in asi_lineas), 2)
        total_deb = round(sum(l['debito'] for l in lineas), 2)
        folio = f"{f.get('prefijo', '')}{f.get('folio', '')}"
        estado = 'OK' if rr['sin_cuenta'] == 0 else f"{rr['sin_cuenta']} ítem(s) sin cuenta"
        if p is None:
            estado = 'NUEVO - asignar cuenta'
        asientos.append({
            'fecha': f.get('fecha', ''), 'tipo': 'FE', 'factura': folio,
            'nit': f['nit'], 'proveedor': (rr['nombre'] or '')[:40],
            'total': round(total_deb, 2), 'iva': round(sum(it['iva'] for it in rr['items']), 2),
            'origen': 'XML', 'items': rr['items'], 'lineas': lineas,
            'prov_cuenta': rr['cxp'], 'credito': credito, 'estado': estado,
            'comprobante': '00003', 'dividida': len(lineas) > 2,
            'cuadra': abs(total_deb - credito) < 1,
            'sin_cuenta': rr['sin_cuenta'],
        })
    resumen = {
        'facturas': len(asientos), 'total': round(sum(a['total'] for a in asientos), 2),
        'ok': sum(1 for a in asientos if a['estado'] == 'OK'),
        'con_xml': len(asientos), 'divididas': sum(1 for a in asientos if a['dividida']),
        'nuevos': sum(1 for a in asientos if 'NUEVO' in a['estado']),
        'revisar': sum(1 for a in asientos if 'sin cuenta' in a['estado']),
        'descuadres': sum(1 for a in asientos if not a['cuadra']),
        'items_total': sum(len(a['items']) for a in asientos),
    }
    return asientos, resumen


def clasificar(ruta_dian, perfiles, facturas_xml=None, reglas=None):
    """Devuelve (lista_asientos, resumen)."""
    facturas_xml = facturas_xml or {}
    df = pd.read_excel(ruta_dian, header=0, dtype=str)
    rec = df[df['Grupo'].astype(str).str.contains('ecib', na=False)].copy()
    for c in ['IVA', 'INC', 'Total']:
        rec[c] = _num(rec[c]).fillna(0)
    rec['nit'] = rec['NIT Emisor'].apply(limpiar_nit)
    rec['fe'] = pd.to_datetime(rec['Fecha Emisión'], format='%d-%m-%Y', errors='coerce')
    rec = rec.sort_values('fe')

    asientos = []
    for _, r in rec.iterrows():
        nota = 'rédito' in str(r['Tipo de documento'])
        signo = -1 if nota else 1
        T, V = r['Total'], r['IVA']
        nit = r['nit']
        p = perfiles.get(nit)
        folio = f"{r['Prefijo'] or ''}{r['Folio'] or ''}"
        clave = nit + '_' + str(r['Folio'] or '')
        fac = facturas_xml.get(clave) or facturas_xml.get(str(r['Folio'] or ''))

        lineas = []
        estado = 'OK'
        dividida = False

        if p is None:
            lineas.append({'cuenta': '', 'debito': round((T - V) * signo, 2)})
            if V:
                lineas.append({'cuenta': '', 'debito': round(V * signo, 2)})
            prov_cuenta = '220501'
            estado = 'NUEVO - asignar cuenta'
        elif fac and len(fac['lineas']) > 1 and len(p['cuentas']) > 1:
            dividida = True
            acc = {}
            for ln in fac['lineas']:
                cta = cuenta_por_concepto(ln['concepto'], reglas)
                if cta is None:
                    cta = p['grav_base'] or p['nograv']
                base_total = ln['base'] + ln['iva']
                acc[cta] = acc.get(cta, 0) + base_total
            for cta, val in acc.items():
                lineas.append({'cuenta': cta, 'debito': round(val * signo, 2)})
            prov_cuenta = p['prov']
            estado = 'Dividida por XML'
        else:
            base = T - V
            if V > 0:
                lineas.append({'cuenta': p['grav_base'] or p['nograv'], 'debito': round(base * signo, 2)})
                lineas.append({'cuenta': p['iva'], 'debito': round(V * signo, 2)})
            else:
                lineas.append({'cuenta': p['nograv'] or p['grav_base'], 'debito': round(T * signo, 2)})
            prov_cuenta = p['prov']
            if p['ret']:
                estado = 'Revisar honorario/retención'
            elif len(p['cuentas']) > 1 and not fac:
                estado = 'Mixto sin XML - revisar'

        total_deb = round(sum(l['debito'] for l in lineas), 2)
        asientos.append({
            'fecha': r['fe'].strftime('%d/%m/%Y') if pd.notna(r['fe']) else '',
            'tipo': 'NC' if nota else 'FE',
            'factura': folio, 'nit': nit,
            'proveedor': str(r['Nombre Emisor'])[:40],
            'total': round(T * signo, 2), 'iva': round(V * signo, 2),
            'lineas': lineas, 'prov_cuenta': prov_cuenta,
            'credito': total_deb, 'estado': estado, 'dividida': dividida,
            'cuadra': abs(total_deb - round(T * signo, 2)) < 1,
        })

    resumen = {
        'facturas': len(asientos),
        'total': round(sum(a['total'] for a in asientos), 2),
        'ok': sum(1 for a in asientos if a['estado'] == 'OK'),
        'divididas': sum(1 for a in asientos if a['dividida']),
        'nuevos': sum(1 for a in asientos if 'NUEVO' in a['estado']),
        'revisar': sum(1 for a in asientos if 'Revisar' in a['estado'] or 'sin XML' in a['estado']),
        'descuadres': sum(1 for a in asientos if not a['cuadra']),
    }
    return asientos, resumen
