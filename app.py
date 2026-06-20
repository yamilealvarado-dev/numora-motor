import os
import tempfile
import traceback
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from engine.aprender import aprender_perfiles, cargar_puc
from engine.xml_parser import leer_zip
from engine.clasificar import clasificar
from engine.exportar import generar_excel, generar_txt

app = Flask(__name__)
CORS(app)  # permite que Lovable (u otro frontend) llame al motor
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024  # 60 MB
SALIDA = tempfile.gettempdir()
_cache = {}  # token -> {asientos, resumen, puc}

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


@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        aux = request.files.get('auxiliar')
        dian = request.files.get('dian')
        xmlzip = request.files.get('xml')
        puc_f = request.files.get('puc')
        if not aux or not dian:
            return jsonify({'error': 'Faltan el auxiliar del año anterior y el reporte DIAN.'}), 400

        tmp = tempfile.mkdtemp()
        aux_p = os.path.join(tmp, 'aux.xlsx'); aux.save(aux_p)
        dian_p = os.path.join(tmp, 'dian.xlsx'); dian.save(dian_p)
        puc = {}
        if puc_f:
            puc_p = os.path.join(tmp, 'puc.xlsx'); puc_f.save(puc_p)
            puc = cargar_puc(puc_p)
        facturas = {}
        if xmlzip:
            zp = os.path.join(tmp, 'xml.zip'); xmlzip.save(zp)
            facturas = leer_zip(zp)

        perfiles = aprender_perfiles(aux_p)
        asientos, resumen = clasificar(dian_p, perfiles, facturas)

        token = next(tempfile._get_candidate_names())
        _cache[token] = {'asientos': asientos, 'resumen': resumen, 'puc': puc}

        # resumen por mes para que el frontend muestre las tarjetas de meses
        meses = {}
        for a in asientos:
            m = _mes_de(a)
            d = meses.setdefault(m, {'mes': m, 'nombre': MESES.get(m, m),
                                     'facturas': 0, 'total': 0})
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
    """Recibe los asientos YA EDITADOS por el usuario y genera el archivo final.
    La partida doble se recalcula con las ediciones."""
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
