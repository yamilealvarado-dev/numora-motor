# Numora — Motor de Compras

App web que contabiliza compras automáticamente: aprende del auxiliar del año anterior,
clasifica las facturas del reporte DIAN, divide por concepto usando los XML, arma la
partida doble y genera el Excel de revisión + el TXT para ContaI.

## Qué hace
1. Subes el **auxiliar del año anterior** (.xlsx) → aprende la cuenta de cada proveedor.
2. Subes el **reporte DIAN** (.xlsx) → las compras a contabilizar.
3. (Opcional) Subes el **ZIP de XML** → divide las facturas mixtas por concepto.
4. (Opcional) Subes el **PUC** → muestra el nombre de cada cuenta.
5. Descargas el **Excel de revisión** y el **TXT para ContaI**.

## Cómo subirlo a Render (gratis)

### 1. Crear cuenta en GitHub (gratis)
- Entra a github.com → Sign up.
- Crea un repositorio nuevo (botón "New") llamado, por ejemplo, `numora-motor`.
- En el repo, botón "Add file" → "Upload files" → arrastra TODO el contenido de esta
  carpeta (app.py, la carpeta engine, la carpeta templates, requirements.txt, Procfile,
  render.yaml) → "Commit changes".

### 2. Crear cuenta en Render (gratis)
- Entra a render.com → Sign up (puedes entrar con tu cuenta de GitHub).
- Botón "New +" → "Web Service".
- Conecta tu repositorio `numora-motor`.
- Render detecta Python solo. Deja los valores por defecto:
  - Build command: `pip install -r requirements.txt`
  - Start command: `gunicorn app:app`
  - Plan: **Free** (o Starter $7/mes si lo quieres siempre prendido).
- Botón "Create Web Service".

### 3. Listo
- En unos minutos Render te da una dirección tipo `https://numora-motor.onrender.com`.
- Ábrela y empieza a subir tus archivos.

## Probarlo en tu computador (opcional)
```
pip install -r requirements.txt
python app.py
```
Luego abre http://localhost:5000

## Editar las reglas de palabras clave
El archivo `engine/reglas.py` tiene las palabras que rutean cada línea a su cuenta
(aseo, cafetería, etc.). Ahí se agregan los conceptos típicos de cada empresa.
