# Red Bibliométrica

Aplicación Flask para análisis de redes bibliométricas basada en Excel.

## Archivos importantes

- `app.py` - backend de Flask
- `templates/index.html` - frontend
- `requirements.txt` - dependencias Python
- `Procfile` - comando de arranque para Render
- `render.yaml` - configuración de despliegue opcional para Render

## Despliegue en Render

1. Sube este proyecto a un repositorio de GitHub.
2. En Render, crea un nuevo servicio web.
3. Conecta tu repositorio.
4. Configura:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Despliega.

## Uso del QR

Render te dará una URL pública como `https://<nombre>.onrender.com`.

- Genera un QR con esa URL.
- Al escanearlo, abrirá directamente la aplicación.

## Notas

- El archivo `ARCHIVOS.xlsx` se usa como fuente de datos inicial.
- El archivo subido se guarda como `uploaded_data.xlsx`.
- En entornos como Render, los archivos subidos son temporales y pueden perderse en reinicios.
