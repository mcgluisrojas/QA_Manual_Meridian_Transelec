# API de Planilla de Revisión QA

Este repositorio incluye una API ligera escrita solo con la librería estándar
de Python para exponer los datos del archivo `QA_Planilla_Revision_Manual_Meridian_Transelec.xlsx`.

## Requisitos

- Python 3.11 o superior.
- No se necesitan dependencias externas ni instalación adicional.

## Uso

1. Desde la raíz del repositorio ejecuta:

   ```bash
   python api_server.py --host 0.0.0.0 --port 8000
   ```

2. Endpoints disponibles:
   - `GET /health`: comprueba que el servidor esté en ejecución.
   - `GET /api/tests`: devuelve los encabezados y filas de la hoja 1 del archivo XLSX.

Ejemplo de respuesta de `/api/tests` (la planilla incluida solo contiene los
encabezados en la primera hoja, por lo que la lista de filas puede estar vacía):

```json
{
  "headers": ["ID", "Test Name", "Descripción", "Resultado Esperado", "Estado", "Evidencia"],
  "rows": []
}
```

## Estructura

- `excel_reader.py`: utilidades para leer la planilla sin dependencias externas.
- `api_server.py`: servidor HTTP básico con los endpoints descritos.
- `QA_Planilla_Revision_Manual_Meridian_Transelec.xlsx`: archivo de datos leído por la API.

## Pruebas

Ejecuta las pruebas unitarias con:

```bash
python -m unittest
```
