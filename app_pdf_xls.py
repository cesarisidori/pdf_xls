import streamlit as st
import pdfplumber
import pandas as pd
import io

# Configuración de la página de Streamlit
st.set_page_config(page_title="Extractor de PDFs a Excel", page_icon="📊", layout="wide")

st.title("📊 Extractor de Tablas PDF a Excel (versión base)")
st.write("Sube tu archivo PDF para procesarlo con tu script consolidarlo en un Excel.")

# 1. Selector de archivos (reemplaza las rutas fijas en disco)
archivo_subido = st.file_uploader("Elige un archivo PDF", type=["pdf"])

if archivo_subido is not None:
    # Abrimos el PDF temporalmente para saber cuántas páginas tiene
    with pdfplumber.open(archivo_subido) as pdf:
        total_paginas = len(pdf.pages)
    
    st.success(f"📄 Archivo cargado con éxito. Total de páginas: {total_paginas}")
    
    # 2. Selector de página de inicio (agregado por comodidad para saltar portadas)
    pagina_inicio = st.number_input(
        "¿A partir de qué página deseas empezar a escanear?", 
        min_value=1, 
        max_value=total_paginas, 
        value=1,  # Por defecto desde la 1, tal como tu script original
        help="Las páginas anteriores a esta serán ignoradas."
    )
    
    # 3. Botón para iniciar el procesamiento
    if st.button("🚀 Iniciar Extracción"):
        lista_tablas = []
        tablas_encontradas = 0
        columnas_referencia = None
        
        # Elementos visuales de Streamlit para el progreso
        barra_progreso = st.progress(0)
        texto_estado = st.empty()
        
        # Abrimos el PDF (Manejo en memoria que le gusta a Streamlit)
        with pdfplumber.open(archivo_subido) as pdf:
            paginas_a_procesar = pdf.pages[pagina_inicio - 1:]
            total_a_procesar = len(paginas_a_procesar)
            
            for idx, pagina in enumerate(paginas_a_procesar):
                # Actualizar barra de progreso visual
                porcentaje = int((idx + 1) / total_a_procesar * 100)
                barra_progreso.progress(porcentaje)
                texto_estado.text(f"Procesando página {pagina_inicio + idx} de {total_paginas}...")
                
                # TU LÓGICA ORIGINAL INMUTABLE:
                tablas_pagina = pagina.extract_tables()
                
                for tabla in tablas_pagina:
                    if not tabla or len(tabla) < 2:  # Salta si la tabla está vacía
                        continue
                    
                    tablas_encontradas += 1
                    
                    # Convertimos la matriz de texto a DataFrame usando la primera fila como título
                    df_tabla = pd.DataFrame(tabla[1:], columns=tabla[0])
                    
                    # --- LIMPIEZA DINÁMICA DE ENCABEZADOS REPETIDOS (Tu lógica) ---
                    if not df_tabla.empty:
                        columnas_actuales = [str(c).strip().lower() for c in df_tabla.columns]
                        primera_fila = [str(x).strip().lower() for x in df_tabla.iloc[0].values]
                        
                        if primera_fila == columnas_actuales:
                            df_tabla = df_tabla.iloc[1:].reset_index(drop=True)
                    
                    if df_tabla.empty:
                        continue
                    
                    # --- ESTANDARIZACIÓN GENERALISTA (Tu lógica) ---
                    if columnas_referencia is None:
                        columnas_referencia = df_tabla.columns
                    else:
                        if len(df_tabla.columns) == len(columnas_referencia):
                            df_tabla.columns = columnas_referencia
                    
                    lista_tablas.append(df_tabla)
        
        # Limpieza de los indicadores de progreso al terminar
        texto_estado.empty()
        barra_progreso.empty()
        
        # 4. Consolidación y Descarga
        if lista_tablas:
            # Tu concat original
            df_final = pd.concat(lista_tablas, ignore_index=True)
            
            st.balloons()
            st.success(f"✅ ¡Proceso completado! Se consolidaron {tablas_encontradas} tablas en {len(df_final)} filas.")
            
            # Vista previa para asegurarte de que se vea igual que antes
            #st.subheader("👀 Vista previa de los datos:")
            #st.dataframe(df_final.head(20), use_container_width=True)
            
            # Convertimos a Excel usando BytesIO para que Streamlit lo maneje en el navegador
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final.to_excel(writer, sheet_name="Datos_Consolidados", index=False)
            datos_excel = output.getvalue()
            
            # Botón de descarga nativo de Streamlit
            st.download_button(
                label="📥 Descargar Archivo Excel Consolidado",
                data=datos_excel,
                file_name="Datos_Consolidados_PDF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No se detectaron tablas en el documento.")
