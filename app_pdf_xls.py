import streamlit as st
import pdfplumber
import pandas as pd
import io

# Configuración de la página de Streamlit
st.set_page_config(page_title="Extractor de PDFs a Excel", page_icon="📊", layout="wide")

st.title("📊 Extractor Inteligente de Tablas (PDF a Excel)")
st.write("Sube tu archivo PDF, selecciona desde qué página comenzar el escaneo y consolida todo en un único Excel limpio.")

# 1. Selector de archivos
archivo_subido = st.file_uploader("Elige un archivo PDF", type=["pdf"])

if archivo_subido is not None:
    # Abrimos el PDF temporalmente para saber cuántas páginas tiene
    with pdfplumber.open(archivo_subido) as pdf:
        total_paginas = len(pdf.pages)
    
    st.success(f"📄 Archivo cargado con éxito. Total de páginas: {total_paginas}")
    
    # 2. Configuración por parte del usuario (Widgets de Streamlit)
    col1, col2 = st.columns(2)
    with col1:
        pagina_inicio = st.number_input(
            "¿A partir de qué página deseas empezar a escanear?", 
            min_value=1, 
            max_value=total_paginas, 
            value=3, # Valor por defecto (página 3, saltándose portadas/índices)
            help="Las páginas anteriores a esta serán ignoradas por completo."
        )
    
    # 3. Botón para iniciar el procesamiento
    if st.button("🚀 Iniciar Extracción y Consolidación"):
        lista_tablas = []
        tablas_encontradas = 0
        columnas_referencia = None
        
        # Crear una barra de progreso visual en la web
        barra_progreso = st.progress(0)
        texto_estado = st.empty()
        
        # Volvemos a abrir el archivo para procesarlo página por página
        with pdfplumber.open(archivo_subido) as pdf:
            # Filtramos las páginas según la selección del usuario (ajustando el índice que empieza en 0)
            paginas_a_procesar = pdf.pages[pagina_inicio - 1:]
            total_a_procesar = len(paginas_a_procesar)
            
            for idx, pagina in enumerate(paginas_a_procesar):
                # Actualizar progreso en la interfaz web
                porcentaje = int((idx + 1) / total_a_procesar * 100)
                barra_progreso.progress(porcentaje)
                texto_estado.text(f"Procesando página {pagina_inicio + idx} de {total_paginas}...")
                
                # Extraer tablas de la página actual
                tablas_pagina = pagina.extract_tables()
                
                for tabla in tablas_pagina:
                    if not tabla or len(tabla) < 2:
                        continue
                    
                    tablas_encontradas += 1
                    df_tabla = pd.DataFrame(tabla[1:], columns=tabla[0])
                    
                    # Limpieza dinámica de encabezados repetidos por saltos de página
                    if not df_tabla.empty:
                        columnas_actuales = [str(c).strip().lower() for c in df_tabla.columns]
                        primera_fila = [str(x).strip().lower() for x in df_tabla.iloc[0].values]
                        
                        if primera_fila == columnas_actuales:
                            df_tabla = df_tabla.iloc[1:].reset_index(drop=True)
                    
                    if df_tabla.empty:
                        continue
                    
                    # Estandarización generalista de nombres de columnas coincidentes
                    if columnas_referencia is None:
                        columnas_referencia = df_tabla.columns
                    else:
                        if len(df_tabla.columns) == len(columnas_referencia):
                            df_tabla.columns = columnas_referencia
                    
                    lista_tablas.append(df_tabla)
        
        # Limpieza de textos y barra al finalizar
        texto_estado.empty()
        barra_progreso.empty()
        
        # 4. Mostrar resultados y permitir la descarga
        if lista_tablas:
            df_final = pd.concat(lista_tablas, ignore_index=True)
            
            # 1. Limpieza de columnas fantasma conocidas
            columnas_a_borrar = r'\(hh:mm\)|al año'
            df_final = df_final.loc[:, ~df_final.columns.str.contains(columnas_a_borrar, case=False, na=False)]
            
            # 2. LIMPIEZA DE COLUMNAS VACÍAS O SIN NOMBRE
            df_final = df_final.loc[:, df_final.columns.notna()]
            df_final = df_final.loc[:, df_final.columns != '']
            df_final = df_final.loc[:, ~df_final.columns.str.contains('^Unnamed', case=False, na=False)]
            
            # 3. Forzar a que todos los nombres de columna sean únicos (Solución definitiva para PyArrow)
            columnas_unicas = []
            conteos = {}
            for col in df_final.columns:
                col_str = str(col).strip()
                if col_str in conteos:
                    conteos[col_str] += 1
                    columnas_unicas.append(f"{col_str}_{conteos[col_str]}")
                else:
                    conteos[col_str] = 0
                    columnas_unicas.append(col_str)
            df_final.columns = columnas_unicas
            
            st.balloons() 
            st.success(f"✅ ¡Proceso completado! Se consolidaron {tablas_encontradas} tablas en un total de {len(df_final)} filas.")
            
            # Mostrar la vista previa de los datos
            st.subheader("👀 Vista previa de los datos consolidados:")
            st.dataframe(df_final.head(20), use_container_width=True)
            
            # 5. Botón de descarga de Excel nativo de Streamlit
            st.download_button(
                label="📥 Descargar Archivo Excel Consolidado",
                data=datos_excel,
                file_name="Reporte_Consolidado_PDF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No se detectaron tablas en las páginas seleccionadas. Intenta cambiando la página de inicio.")
