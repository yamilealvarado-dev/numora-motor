"""Toma el reporte DIAN + perfiles + (opcional) líneas de XML y arma el asiento de cada compra."""
import re
import pandas as pd
from .aprender import limpiar_nit
from .reglas import cuenta_por_concepto


def _num(x):
    return pd.to_numeric(x, errors='coerce')


def clasificar(ruta_dian, perfiles, facturas_xml=None, reglas=None):
    """Devuelve (lista_asientos, resumen).
    Cada asiento: {fecha, tipo, factura, nit, proveedor, total, lineas:[{cuenta,nombre,debito}],
                   prov_cuenta, iva, estado, dividida}."""
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
            # proveedor nuevo
            lineas.append({'cuenta': '', 'debito': round((T - V) * signo, 2)})
            if V:
                lineas.append({'cuenta': p['iva'] if p else '', 'debito': round(V * signo, 2)})
            prov_cuenta = '220501'
            estado = 'NUEVO - asignar cuenta'
        elif fac and len(fac['lineas']) > 1 and len(p['cuentas']) > 1:
            # proveedor mixto + tenemos XML -> dividir por línea
            dividida = True
            acc = {}
            for ln in fac['lineas']:
                cta = cuenta_por_concepto(ln['concepto'], reglas)
                if cta is None:
                    cta = p['grav_base'] or p['nograv']  # cuenta principal del proveedor
                base_total = ln['base'] + ln['iva']  # IVA al costo
                acc[cta] = acc.get(cta, 0) + base_total
            for cta, val in acc.items():
                lineas.append({'cuenta': cta, 'debito': round(val * signo, 2)})
            prov_cuenta = p['prov']
            estado = 'Dividida por XML'
        else:
            # proveedor conocido, una sola cuenta (o sin XML)
            base = T - V
            if V > 0:
                lineas.append({'cuenta': p['grav_base'] or p['nograv'], 'debito': round(base * signo, 2)})
                lineas.append({'cuenta': p['iva'], 'debito': round(V * signo, 2)})
            else:
                lineas.append({'cuenta': p['nograv'] or p['grav_base'], 'debito': round(T * signo, 2)})
            prov_cuenta = p['prov']
            if p['ret']:
                estado = 'Revisar honorario/retención'
            elif nit in perfiles and len(p['cuentas']) > 1 and not fac:
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
