import os
import tempfile
import traceback
import requests
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from engine.aprender import aprender_perfiles, aprender_iva_pares, cargar_puc
from engine.xml_parser import leer_zip
from engine.clasificar import clasificar, clasificar_dian_items, clasificar_solo_xml
from engine.exportar import generar_excel, generar_txt

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024
SALIDA = tempfile.gettempdir()
_cache = {}

MESES = {'01': 'Enero', '02': 'Febrero', '03': 'Marzo', '04': 'Abril', '05': 'Mayo',
         '06': 'Junio', '07': 'Julio', '08': 'Agosto', '09': 'Septiembre',
         '10': 'Octubre', '11': 'Noviembre', '12': 'Diciembre'}


def _mes_de(a):
    try:
        return a['fecha'].split('/')[1]
    except Exception:
        return '00'


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/alegra-test')
def alegra_test():
    """Prueba la conexión con Alegra y muestra cómo vienen los datos."""
    from engine import alegra
    if not alegra.hay_credenciales():
        return jsonify({'ok': False, 'mensaje': 'Faltan ALEGRA_EMAIL o ALEGRA_TOKEN en Render.'}), 400
    desde = request.args.get('desde')
    hasta = request.args.get('hasta')
    try:
        return jsonify(alegra.diagnostico(desde, hasta))
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else '?'
        msg = ('Credenciales incorrectas (401)' if code == 401 else
               'Tu plan de Alegra no permite API — se requiere plan Plus (403)' if code == 403 else
               f'Error HTTP {code}')
        return jsonify({'ok': False, 'mensaje': msg, 'detalle': str(e)}), 200
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'mensaje': 'Error de conexión', 'detalle': str(e)}), 200


@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        aux = request.files.get('auxiliar')
        dian = request.files.get('dian')
        xmlzip = request.files.get('xml')
        puc_f = request.files.get('puc')
        if not aux:
            return jsonify({'error': 'Falta el auxiliar (para aprender las cuentas).'}), 400
        if not dian and not xmlzip:
            return jsonify({'error': 'Sube el reporte DIAN o el ZIP de XML (al menos uno).'}), 400

        tmp = tempfile.mkdtemp()
        aux_p = os.path.join(tmp, 'aux.xlsx'); aux.save(aux_p)
        puc = {}
        if puc_f:
            puc_p = os.path.join(tmp, 'puc.xlsx'); puc_f.save(puc_p)
            puc = cargar_puc(puc_p)
        facturas = {}
        if xmlzip:
            zp = os.path.join(tmp, 'xml.zip'); xmlzip.save(zp)
            facturas = leer_zip(zp)

        perfiles = aprender_perfiles(aux_p)
        iva_pares = aprender_iva_pares(aux_p)
        if dian:
            dian_p = os.path.join(tmp, 'dian.xlsx'); dian.save(dian_p)
            asientos, resumen = clasificar_dian_items(dian_p, perfiles, iva_pares, facturas)
        else:
            asientos, resumen = clasificar_solo_xml(facturas, perfiles, iva_pares)

        token = next(tempfile._get_candidate_names())
        _cache[token] = {'asientos': asientos, 'resumen': resumen, 'puc': puc}

        meses = {}
        for a in asientos:
            m = _mes_de(a)
            d = meses.setdefault(m, {'mes': m, 'nombre': MESES.get(m, m), 'facturas': 0, 'total': 0})
            d['facturas'] += 1
            d['total'] += a['total']
        meses = sorted(meses.values(), key=lambda x: x['mes'])

        return jsonify({'token': token, 'resumen': resumen, 'meses': meses,
                        'asientos': asientos, 'aprendidos': len(perfiles)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/generar', methods=['POST'])
def generar():
    """Recibe los asientos YA EDITADOS y genera el archivo final."""
    try:
        data = request.get_json(force=True)
        token = data.get('token')
        asientos = data.get('asientos', [])
        item = _cache.get(token, {})
        puc = item.get('puc', {})
        for a in asientos:
            a['credito'] = round(sum(l['debito'] for l in a['lineas'] if l.get('cuenta')), 2)
            a['cuadra'] = abs(a['credito'] - a['total']) < 1
        resumen = {
            'facturas': len(asientos),
            'total': round(sum(a['total'] for a in asientos), 2),
            'ok': sum(1 for a in asientos if a.get('estado') == 'OK'),
            'divididas': sum(1 for a in asientos if a.get('dividida')),
            'nuevos': sum(1 for a in asientos if 'NUEVO' in str(a.get('estado'))),
            'revisar': sum(1 for a in asientos if 'Revisar' in str(a.get('estado'))),
            'descuadres': sum(1 for a in asientos if not a.get('cuadra')),
            'sin_cuenta': sum(1 for a in asientos if not any(l.get('cuenta') for l in a['lineas'])),
        }
        _cache[token] = {'asientos': asientos, 'resumen': resumen, 'puc': puc}
        return jsonify({'token': token, 'resumen': resumen})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/auditar', methods=['POST'])
def auditar_endpoint():
    """Cruza el auxiliar contabilizado contra el reporte DIAN."""
    try:
        from engine.auditar import auditar, generar_excel_auditoria
        aux = request.files.get('auxiliar')
        dian = request.files.get('dian')
        comp = request.form.get('comprobante', '00003')
        if not aux or not dian:
            return jsonify({'error': 'Sube el auxiliar y el reporte DIAN.'}), 400
        tmp = tempfile.mkdtemp()
        aux_p = os.path.join(tmp, 'aux.xlsx'); aux.save(aux_p)
        dian_p = os.path.join(tmp, 'dian.xlsx'); dian.save(dian_p)
        audit = auditar(aux_p, dian_p, comp)
        token = next(tempfile._get_candidate_names())
        path = os.path.join(SALIDA, f'Auditoria_{token}.xlsx')
        generar_excel_auditoria(audit, path)
        _cache['audit_' + token] = path
        audit['token'] = token
        return jsonify(audit)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/descargar-auditoria/<token>')
def descargar_auditoria(token):
    path = _cache.get('audit_' + token)
    if not path or not os.path.exists(path):
        return 'No encontrado', 404
    return send_file(path, as_attachment=True, download_name='Auditoria_compras.xlsx')


@app.route('/descargar/<kind>/<token>')
@app.route('/descargar/<kind>/<token>/<mes>')
def descargar(kind, token, mes=None):
    item = _cache.get(token)
    if not item:
        return 'No encontrado', 404
    asientos = item['asientos']
    if mes:
        asientos = [a for a in asientos if _mes_de(a) == mes]
    sufijo = f'_{MESES.get(mes, mes)}' if mes else ''
    if kind == 'excel':
        path = os.path.join(SALIDA, f'Revision{sufijo}_{token}.xlsx')
        generar_excel(asientos, item['resumen'], item['puc'], path,
                      f'Revisión de compras{(" - " + MESES.get(mes, mes)) if mes else ""}')
    else:
        path = os.path.join(SALIDA, f'Compras{sufijo}_{token}.txt')
        generar_txt(asientos, path)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
