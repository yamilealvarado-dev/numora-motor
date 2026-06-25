"""Aprende, por proveedor, en qué cuentas contabiliza las compras (desde el auxiliar del año anterior o meses ya contabilizados)."""
import re
import pandas as pd
from collections import defaultdict, Counter


def limpiar_nit(x):
    if pd.isna(x):
        return ''
    return re.sub(r'\D', '', str(x).split('-')[0])


def _es_iva(nombre):
    n = str(nombre).upper()
    return 'IVA' in n and ('MAYOR VALOR' in n or 'DESCONT' in n or 'ASUMIDO' in n)


def _es_inc(nombre):
    return 'CONSUMO' in str(nombre).upper()


def _es_proveedor(cuenta):
    return cuenta.startswith('2205') or cuenta.startswith('2335') or cuenta.startswith('2336')


def _es_retencion(cuenta):
    return cuenta.startswith('2365') or cuenta.startswith('2367') or cuenta.startswith('2368')


def aprender_perfiles(ruta_auxiliar, comprobante_compra='00003'):
    """Devuelve {nit: {grav_base, iva, nograv, prov, inc, ret, cuentas, n, nombre}}."""
    aux = pd.read_excel(ruta_auxiliar, sheet_name='Datos', header=2, dtype=str)
    for c in ['Débitos', 'Créditos']:
        aux[c] = pd.to_numeric(aux[c], errors='coerce').fillna(0)
    aux['nit'] = aux['NIT'].apply(limpiar_nit)
    comp = aux[aux['Comprobante'] == comprobante_compra].copy()

    perfil = defaultdict(lambda: {'grav_base': Counter(), 'iva': Counter(), 'nograv': Counter(),
                                  'prov': Counter(), 'inc': Counter(), 'ret': Counter(),
                                  'n': 0, 'nombre': ''})
    for doc, a in comp.groupby('Documento'):
        if not doc or pd.isna(doc):
            continue
        pl = a[a['Cuenta'].apply(_es_proveedor) & (a['Créditos'] > 0)]
        if len(pl) == 0:
            continue
        nit = pl.iloc[0]['nit']
        if len(nit) < 7:
            continue
        p = perfil[nit]
        p['n'] += 1
        p['nombre'] = pl.iloc[0]['Nombre NIT']
        p['prov'][pl.iloc[0]['Cuenta']] += 1
        debs = a[a['Débitos'] > 0]
        iva_l = debs[debs['Nombre Cuenta'].apply(_es_iva)]
        inc_l = debs[debs['Nombre Cuenta'].apply(_es_inc)]
        base_l = debs[~debs.index.isin(iva_l.index) & ~debs.index.isin(inc_l.index)]
        if len(iva_l) > 0:
            p['iva'][iva_l.iloc[0]['Cuenta']] += 1
            for _, b in base_l.iterrows():
                p['grav_base'][b['Cuenta']] += 1
        else:
            for _, b in base_l.iterrows():
                p['nograv'][b['Cuenta']] += 1
        for _, b in inc_l.iterrows():
            p['inc'][b['Cuenta']] += 1
        for _, r in a[a['Cuenta'].apply(_es_retencion) & (a['Créditos'] > 0)].iterrows():
            p['ret'][r['Cuenta']] += 1

    def top(c, d=None):
        return c.most_common(1)[0][0] if c else d

    return {nit: {'grav_base': top(p['grav_base']), 'iva': top(p['iva']),
                  'nograv': top(p['nograv']), 'prov': top(p['prov'], '220501'),
                  'inc': top(p['inc']), 'ret': top(p['ret']),
                  'cuentas': sorted(set(list(p['grav_base']) + list(p['nograv']))),
                  'n': p['n'], 'nombre': p['nombre']}
            for nit, p in perfil.items()}


def aprender_iva_pares(ruta_auxiliar, comprobante_compra='00003'):
    """Aprende la pareja {cuenta_costo: cuenta_iva} desde el auxiliar:
    en cada asiento de compra, empareja cada cuenta base con la cuenta de IVA que aparece junto."""
    aux = pd.read_excel(ruta_auxiliar, sheet_name='Datos', header=2, dtype=str)
    for c in ['Débitos', 'Créditos']:
        aux[c] = pd.to_numeric(aux[c], errors='coerce').fillna(0)
    comp = aux[aux['Comprobante'] == comprobante_compra]
    pares = defaultdict(Counter)
    for doc, a in comp.groupby('Documento'):
        debs = a[a['Débitos'] > 0]
        iva_l = debs[debs['Nombre Cuenta'].apply(_es_iva)]
        base_l = debs[~debs.index.isin(iva_l.index)]
        if len(iva_l) > 0:
            iva_acct = iva_l.iloc[0]['Cuenta']
            for _, b in base_l.iterrows():
                pares[b['Cuenta']][iva_acct] += 1
    return {c.strip(): cc.most_common(1)[0][0].strip() for c, cc in pares.items()}


def cargar_puc(ruta_puc):
    puc = pd.read_excel(ruta_puc, sheet_name='Datos', header=2, dtype=str)
    puc['c'] = puc['Cuenta'].apply(lambda x: re.sub(r'\D', '', str(x)))
    return dict(zip(puc['c'], puc['Nombre']))


def cargar_puc_meta(ruta_puc):
    """Carga metadatos del PUC por cuenta: tipo (S/B/C), base de retención y centro de costo.
    Sirve para el TXT ContaExport: definir NIT, valor base de retención y centro de costo."""
    df = pd.read_excel(ruta_puc, sheet_name='Datos', header=2, dtype=str)
    meta = {}
    for _, r in df.iterrows():
        cod = str(r.get('Cuenta', '') or '').replace('-', '').strip()
        if not cod:
            continue
        meta[cod] = {
            'tipo': str(r.get('Tipo', '') or '').strip().upper(),
            'baseret': str(r.get('Base Ret.', '') or '').strip(),
            'ccosto': str(r.get('Ccosto', '') or '').strip().upper(),
            'nombre': str(r.get('Nombre', '') or '').strip(),
        }
    return meta
