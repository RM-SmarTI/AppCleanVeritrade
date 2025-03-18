import pandas as pd
import streamlit as st
from io import BytesIO

# Función para procesar los datos
def procesar_datos(veritrade_file, palabras_excluir, palabras_incluir):
    palabras_excluir = [palabra.strip() for palabra in palabras_excluir.split(',') if palabra.strip()]
    palabras_incluir = [palabra.strip() for palabra in palabras_incluir.split(',') if palabra.strip()]
    
    try:
        # Leer el archivo subido correctamente
        veritrade = pd.read_excel(veritrade_file, sheet_name='Veritrade', skiprows=5, engine='openpyxl')
        veritrade.rename(columns={'Descripcion Comercial': 'DComercial'}, inplace=True)
        
        # Filtrar los datos
        veritrade_clean = veritrade[~veritrade.DComercial.str.contains('|'.join(palabras_excluir), case=False)]
        veritrade_excluir = veritrade[veritrade.DComercial.str.contains('|'.join(palabras_excluir), case=False)]

        if palabras_incluir:  # Si hay palabras a incluir
            veritrade_incluir = veritrade_excluir[veritrade_excluir.DComercial.str.contains('|'.join(palabras_incluir), case=False)]
            veritrade_excluir2 = veritrade_excluir[~veritrade_excluir.DComercial.str.contains('|'.join(palabras_incluir), case=False)]
            veritrade_clean2 = pd.concat([veritrade_clean, veritrade_incluir], ignore_index=True)
        else: # Si palabras_incluir está vacío, no filtramos
            veritrade_clean2 = veritrade_clean
            veritrade_excluir2 = veritrade_excluir
        
        # Guardar el archivo procesado en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            veritrade_clean2.to_excel(writer, sheet_name='data_limpia', index=False)
            veritrade_excluir2.to_excel(writer, sheet_name='exclusiones', index=False)
        writer.close()
        output.seek(0)
        
        return output
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los datos: {str(e)}")
        return None

# Configuración de la aplicación Streamlit
st.title("Procesador de Veritrade")
st.write("Sube un archivo Excel de Veritrade, ingresa palabras a excluir y descarga el archivo procesado.")

# Cargar archivo de Veritrade
archivo_veritrade = st.file_uploader("Sube el archivo Excel de Veritrade", type=["xlsx", "xls"])

# Entrada para palabras a incluir
palabras_incluir = st.text_input("Palabras a incluir (separadas por comas)")

# Entrada para palabras a excluir
palabras_excluir = st.text_input("Palabras a excluir (separadas por comas)")

# Botón para procesar datos
if st.button("Procesar Datos"):
    if archivo_veritrade and palabras_excluir:
        archivo_procesado = procesar_datos(archivo_veritrade, palabras_excluir, palabras_incluir)
        if archivo_procesado:
            st.success("Archivo Excel generado con éxito.")
            st.download_button(label="Descargar Archivo Procesado", data=archivo_procesado, file_name="veritrade_procesado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Por favor, sube un archivo y escribe palabras a excluir.")
