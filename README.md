# Validador de Derivados Streamlit

App de Streamlit que une los procesos de **Listados** y **OTC/CSA** en una sola interfaz.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

- `app.py`: interfaz principal.
- `src/listados.py`: proceso de Validacion Derivados Listados.
- `src/otc.py`: proceso de Validaciones OTC CSA.
- `src/pdf_statements.py`: lectura de estados de cuenta PDF.
- `src/excel_utils.py`: utilidades compartidas para Excel.

## Notas importantes

Algunas plantillas dependen de formulas de Excel. En Windows, si esta disponible Microsoft Excel, la app intenta recalcular automaticamente con COM. En Mac/Linux la app conserva las formulas, pero los valores calculados dependen de que la plantilla haya sido guardada previamente con resultados actualizados.
