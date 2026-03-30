# Acomodador de Layout - Streamlit

App para homologar varios archivos de Excel contra un layout base.

## Qué hace
- Pide primero el layout vigente.
- Después permite subir múltiples archivos Excel.
- Usa el encabezado del layout como estructura principal.
- Si un archivo no trae una columna del layout, la deja vacía.
- Si un archivo trae columnas que no existen en el layout, las agrega al final.
- Resalta con otro color las columnas extra.
- Deja la fila 2 vacía y comienza a escribir datos desde la fila 3.
- Puede agregar la columna **Archivo origen**.

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue
Sube estos archivos a un repositorio nuevo en GitHub:
- app.py
- requirements.txt
- README.md

Luego en Streamlit Community Cloud:
1. New app
2. Selecciona tu repositorio
3. Branch: main
4. Main file path: app.py
5. Deploy
