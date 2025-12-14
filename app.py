import streamlit as st
import pandas as pd
import io
import xlsxwriter

# Configuración visual de la página
st.set_page_config(page_title="Monitor Limpio CBM", page_icon="📊")

st.title("📊 Limpiador de Monitor - Ventas")
st.markdown("Sube el **Monitor Oficial** y descárgalo listo para Looker.")

# 1. WIDGET DE SUBIDA
uploaded_file = st.file_uploader("Arrastra aquí el archivo Excel (.xlsx)", type=["xlsx"])

# Aquí estaba el error, el 'if' debe ir abajo
if uploaded_file is not None:
    try:
        with st.spinner('Limpiando datos... ⏳'):
            
            # 1. CARGA Y FILTRADO
            # Columnas exactas que definimos
            cols_necesarias = "B,D,E,J,N,O,Q,R,S,T,U,V,W,X,Y,AD"
            
            df = pd.read_excel(
                uploaded_file, 
                sheet_name='Desconsolidacion',
                header=8,
                usecols=cols_necesarias
            )
            
            # Limpiamos espacios en nombres de columnas
            df.columns = df.columns.str.strip()

            # 2. SEPARAR FECHA Y HORA
            if 'FECHA DESCO' in df.columns:
                df['FECHA DESCO'] = pd.to_datetime(df['FECHA DESCO'], errors='coerce')
                df['HORA DESCO'] = df['FECHA DESCO'].dt.time
                df['FECHA DESCO'] = df['FECHA DESCO'].dt.date
            
            st.success("✅ ¡Listo! Datos limpios.")
            
            # Mostramos vista previa
            st.dataframe(df.head(3))
            
            # 3. GENERAR EXCEL
            buffer = io.BytesIO()
            # Usamos xlsxwriter explícitamente
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='DATA_LIMPIA')
                
                # Autoajustar columnas para que se vea bonito
                worksheet = writer.sheets['DATA_LIMPIA']
                for i, col in enumerate(df.columns):
                    # Calculamos el ancho ideal
                    max_len = df[col].astype(str).map(len).max()
                    if pd.isna(max_len): max_len = 0
                    width = max(max_len, len(str(col))) + 2
                    worksheet.set_column(i, i, width)

            buffer.seek(0)
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Excel (.xlsx)",
                data=buffer,
                file_name="Monitor_Limpio_Looker.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"❌ Error: {e}")