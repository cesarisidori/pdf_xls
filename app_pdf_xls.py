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
            
            # Tolerancia visual para números compactos y centrados
            ajustes_tabla = {
                "vertical_strategy": "text", 
                "horizontal_strategy": "text",
                "snap_tolerance": 4,      
                "text_tolerance": 4       
            }
            
            for idx, pagina in enumerate(paginas_a_procesar):
                # Actualizar progreso en la interfaz web
                porcentaje = int((idx + 1) / total_a_procesar * 100)
                barra_progreso.progress(porcentaje)
                texto_estado.text(f"Procesando página {pagina_inicio + idx} de {total_paginas}...")
                
                # Extraer tablas aplicando la configuración especial de tolerancia visual
                tablas_pagina = pagina.extract_tables(table_settings=ajustes_tabla)
                
                for tabla in tablas_pagina:
                    if not tabla or len(tabla) < 2:
                        continue
                    
                    tablas_encontradas += 1
                    
                    # Limpiamos y preparamos encabezados temporales para esta página
                    encabezados = [str(c).strip() if c is not None else "" for c in tabla[0]]
                    
                    # Renombramos al vuelo nombres vacíos para no perder posiciones físicas
                    encabezados_limpios = []
                    for i, enc in enumerate(encabezados):
                        if enc == "" or "unnamed" in enc.lower():
                            encabezados_limpios.append(f"Columna_{i}")
                        else:
                            encabezados_limpios.append(enc)
                    
                    df_tabla = pd.DataFrame(tabla[1:], columns=encabezados_limpios)
                    
                    # Limpieza dinámica de encabezados repetidos por saltos de página
                    if not df_tabla.empty:
                        primera_fila = [str(x).strip() for x in df_tabla.iloc[0].values]
                        if primera_fila == encabezados_limpios:
                            df_tabla = df_tabla.iloc[1:].reset_index(drop=True)
                    
                    if df_tabla.empty:
                        continue
                    
                    # Forzar acople por posición física de las columnas
                    if columnas_referencia is None:
                        columnas_referencia = df_tabla.columns
                    else:
                        if len(df_tabla.columns) == len(columnas_referencia):
                            df_tabla.columns = columnas_referencia
                        elif len(df_tabla.columns) > len(columnas_referencia):
                            df_tabla = df_tabla.iloc[:, :len(columnas_referencia)]
                            df_tabla.columns = columnas_referencia
                    
                    lista_tablas.append(df_tabla)
        
        # Limpieza de textos y barra al finalizar
        texto_estado.empty()
        barra_progreso.empty()
        
        # 4. Mostrar resultados y permitir la descarga
        if lista_tablas:
            df_final = pd.concat(lista_tablas, ignore_index=True)
            
            # 1. Limpieza de columnas basura de texto conocidas al final
            columnas_a_borrar = r'\(hh:mm\)|al año'
            df_final = df_final.loc[:, ~df_final.columns.str.contains(columnas_a_borrar, case=False, na=False)]
            
            # 2. Quitar únicamente columnas que estén COMPLETAMENTE vacías en todo el documento
            df_final = df_final.dropna(how='all', axis=1)
            
            # 3. Forzar a que todos los nombres de columna sean únicos (Para PyArrow en Streamlit)
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
            
            # 🔥 PARTE NUEVA: Conversión inteligente de texto a números reales (ej: '54,023' -> 54.023)
            for col in df_final.columns:
                # Intentamos procesar columnas que tengan palabras clave como "km", "distancia", "origen" o "longitud"
                # O de manera generalizada, columnas donde la mayoría de los valores parezcan números estructurados
                col_lower = col.lower()
                if any(k in col_lower for k in ["km", "distancia", "origen", "longitud", "long", "columna"]):
                    # Eliminamos espacios en blanco, convertimos a string para asegurar el método .str
                    valores_limpios = df_final[col].astype(str).str.strip()
                    
                    # Caso común: Si el número tiene comas como decimales y puntos de miles (ej: 1.234,56 o 54,023)
                    # Para simplificar y estandarizar la extracción de pdfplumber (que suele extraer con coma decimal),
                    # primero quitamos los puntos de miles (si los hay) y cambiamos la coma por el punto decimal de Python.
                    valores_limpios = valores_limpios.str.replace('.', '', regex=False) # Quita puntos de miles
                    valores_limpios = valores_limpios.str.replace(',', '.', regex=False) # Convierte coma decimal a punto
                    
                    # Convertimos a formato numérico de pandas. Si encuentra un texto no numérico, lo deja como NaN (Coerce)
                    df_final[col] = pd.to_numeric(valores_limpios, errors='coerce')
            
            st.balloons() 
            st.success(f"✅ ¡Proceso completado! Se consolidaron {tablas_encontradas} tablas en un total de {len(df_final)} filas.")
            
            # Mostrar la vista previa de los datos
            st.subheader("👀 Vista previa de los datos consolidados:")
            st.dataframe(df_final.head(20), use_container_width=True)
            
            # Conversión a Bytes del Excel en memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
            datos_excel = output.getvalue()
            
            # 5. Botón de descarga de Excel nativo de Streamlit
            st.download_button(
                label="📥 Descargar Archivo Excel Consolidado",
                data=datos_excel,
                file_name="Reporte_Consolidado_PDF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No se detectaron tablas en las páginas seleccionadas. Intenta cambiando la página de inicio.")
