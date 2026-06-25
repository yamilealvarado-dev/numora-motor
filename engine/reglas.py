"""Reglas de palabras clave: rutean cada línea de factura a su cuenta cuando el proveedor es mixto.
Editable: aquí se agregan los conceptos típicos de cada empresa."""

REGLAS_DEFECTO = [
    (['GUANTE', 'JABON', 'JABÓN', 'CLORO', 'DETERGENTE', 'ESCOBA', 'TRAPERO', 'BOLSA',
      'SERVILLETA', 'PAPEL HIGIENICO', 'PAPEL HIGIÉNICO', 'LIMPIA', 'DESINFECT',
      'AROMATIZANTE', 'AMBIENTADOR', 'TOALLA', 'GUANTES'], '519525'),  # Aseo
    (['CAFE', 'CAFÉ', 'AZUCAR', 'AZÚCAR', 'VASO', 'MEZCLADOR', 'AGITADOR',
      'AROMATICA', 'AROMÁTICA'], '519520'),  # Cafetería
    (['HONORARIO', 'ASESORIA', 'ASESORÍA', 'CONTABLE', 'REVISORIA'], '511095'),  # Honorarios
    (['ARRENDAMIENTO', 'ARRIENDO', 'CANON'], '512035'),  # Arrendamientos
    (['ENERGIA', 'ENERGÍA', 'ACUEDUCTO', 'ALCANTARILLADO', 'GAS DOMICILIARIO',
      'TELEFONO', 'TELÉFONO', 'INTERNET'], '513530'),  # Servicios públicos
    (['MANTENIMIENTO', 'REPARACION', 'REPARACIÓN', 'REPUESTO'], '514515'),  # Mantenimiento
]


def cuenta_por_concepto(concepto, reglas=None):
    """Devuelve la cuenta sugerida según palabras clave, o None si no coincide ninguna."""
    reglas = reglas or REGLAS_DEFECTO
    c = (concepto or '').upper()
    for palabras, cuenta in reglas:
        if any(p in c for p in palabras):
            return cuenta
    return None
