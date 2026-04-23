import streamlit as st
import pandas as pd
import time
import random # NUEVO: Importado para las pausas aleatorias
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
# NUEVOS IMPORTS PARA LA NUBE
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# --- FUNCIÓN DE EXTRACCIÓN ---
def extraer_dato_sunat(driver, label, mantener_saltos=False):
    try:
        xpath_str = f"//*[contains(text(), '{label}') and not(*[contains(text(), '{label}')])]"
        elementos = driver.find_elements(By.XPATH, xpath_str)
        if not elementos:
            elementos = driver.find_elements(By.XPATH, f"//*[contains(text(), '{label}')]")
        if not elementos: return "-"

        nodo_label = elementos[0]
        padre = nodo_label.find_element(By.XPATH, "..")
        texto_padre = padre.text.strip()
        valor = texto_padre.replace(nodo_label.text, "").strip()
        if valor.startswith(":"): valor = valor[1:].strip()

        if len(valor) < 2:
            try:
                hermano = nodo_label.find_element(By.XPATH, "following-sibling::*[1]")
                valor = hermano.text.strip()
            except:
                try:
                    hermano_padre = padre.find_element(By.XPATH, "following-sibling::*[1]")
                    valor = hermano_padre.text.strip()
                except: pass

        if not mantener_saltos and valor and "\n" in valor:
            valor = valor.split("\n")[0].strip()
        return valor if valor else "-"
    except:
        return "-"

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Consultor RUC SUNAT", page_icon="🚀")
st.title("🔍 Extractor Masivo SUNAT")

archivo_subido = st.file_uploader("Sube tu Excel con columna 'RUC'", type=["xlsx"])

if archivo_subido:
    df = pd.read_excel(archivo_subido)
    if 'RUC' not in df.columns:
        st.error("❌ Falta la columna 'RUC'")
    else:
        if st.button("🚀 Iniciar Procesamiento"):
            df['RUC'] = df['RUC'].astype(str).str.replace(r'[^\d]', '', regex=True).str.zfill(11)
            
            # --- MODIFICADO: Inicializar TODAS tus columnas si no existen ---
            columnas_extra = ['Razon Social', 'Tipo Contribuyente', 'Tipo de Documento', 'Nombre Comercial', 'Afecto RUS', 'Estado']
            for col in columnas_extra:
                if col not in df.columns:
                    df[col] = '-'

            progreso = st.progress(0)
            status_text = st.empty()

            # --- CONFIGURACIÓN ROBUSTA DE SELENIUM ---
            opciones = webdriver.ChromeOptions()
            opciones.add_argument("--headless=new") 
            opciones.add_argument("--no-sandbox")
            opciones.add_argument("--disable-dev-shm-usage")
            opciones.add_argument("--ignore-certificate-errors")
            opciones.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # NUEVO: Variables para el control de relevos de memoria y errores
            lote_maximo = 150
            filas_procesadas = 0
            hubo_error_fatal = False

            # Usamos un bloque "with" o try/finally para asegurar que el driver se cierre
            try:
                # MODIFICACIÓN PARA LA NUBE: Instalación automática del driver
                servicios = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
                driver = webdriver.Chrome(service=servicios, options=opciones)
                
                url_sunat = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"

                for index, row in df.iterrows():
                    ruc_consulta = row['RUC'].strip()

                    # NUEVO: SISTEMA DE RELEVOS (Libera memoria RAM cada 150 registros)
                    if filas_procesadas > 0 and filas_procesadas % lote_maximo == 0:
                        driver.quit()
                        time.sleep(3) # Respiramos
                        driver = webdriver.Chrome(service=servicios, options=opciones)

                    status_text.info(f"⏳ Procesando: {ruc_consulta} ({index+1}/{len(df)})")
                    
                    reintentos = 3
                    exito = False
                    
                    while reintentos > 0 and not exito:
                        try:
                            driver.get(url_sunat)
                            driver.delete_all_cookies() # NUEVO: Limpia rastros por cada RUC

                            caja_ruc = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "txtRuc")))
                            caja_ruc.clear()
                            caja_ruc.send_keys(ruc_consulta)
                            
                            time.sleep(random.uniform(0.3, 0.8)) # NUEVO: Pausa aleatoria humana
                            driver.find_element(By.ID, "btnAceptar").click()

                            try:
                                WebDriverWait(driver, 3).until(EC.alert_is_present())
                                driver.switch_to.alert.accept()
                                df.at[index, 'Razon Social'] = "RUC INVÁLIDO"
                                exito = True
                                filas_procesadas += 1 # NUEVO: Aumenta contador
                                continue
                            except: pass

                            if len(driver.window_handles) > 1:
                                driver.switch_to.window(driver.window_handles[-1])

                            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Número de RUC')]")))
                            
                            # --- MODIFICADO: Extracción completa de todas tus variables ---
                            ruc_y_razon = extraer_dato_sunat(driver, "Número de RUC")
                            df.at[index, 'Razon Social'] = ruc_y_razon.split(" - ", 1)[1] if " - " in ruc_y_razon else ruc_y_razon
                            df.at[index, 'Tipo Contribuyente'] = extraer_dato_sunat(driver, "Tipo Contribuyente")
                            df.at[index, 'Tipo de Documento'] = extraer_dato_sunat(driver, "Tipo de Documento")
                            df.at[index, 'Nombre Comercial'] = extraer_dato_sunat(driver, "Nombre Comercial")
                            df.at[index, 'Afecto RUS'] = extraer_dato_sunat(driver, "Afecto al Nuevo RUS")
                            df.at[index, 'Estado'] = extraer_dato_sunat(driver, "Estado")
                            
                            exito = True
                            filas_procesadas += 1 # NUEVO: Aumenta contador
                            
                            if len(driver.window_handles) > 1:
                                driver.close()
                                driver.switch_to.window(driver.window_handles[0])

                        except (WebDriverException, TimeoutException):
                            reintentos -= 1
                            status_text.warning(f"⚠️ Reintentando {ruc_consulta}... ({3-reintentos}/3)")
                            time.sleep(2)
                            if reintentos == 0:
                                df.at[index, 'Razon Social'] = "ERROR DE CONEXIÓN"
                                filas_procesadas += 1 # NUEVO: Aumenta contador para no trabar el ciclo

                    progreso.progress((index + 1) / len(df))
                
                driver.quit()
                st.success("✅ ¡Procesamiento terminado!")

            except Exception as e:
                # NUEVO: Captura el error sin destruir los datos avanzados
                st.error(f"❌ Error fatal en el registro {index+1}: {e}")
                st.warning("⚠️ Se interrumpió la consulta, pero puedes descargar los registros que sí se procesaron.")
                hubo_error_fatal = True
                if 'driver' in locals():
                    driver.quit()

            # --- BOTÓN DE DESCARGA CON SPINNER (Movido fuera del try para asegurar descarga parcial) ---
            with st.spinner('📦 Preparando archivo para descarga...'):
                output = BytesIO()
                # Especificamos el motor 'xlsxwriter'
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                data_excel = output.getvalue()

            # Cambia el texto del botón si hubo un corte
            label_boton = "📥 Descargar Resultados Parciales" if hubo_error_fatal else "📥 Descargar Resultados Finales"

            st.download_button(
                label=label_boton, 
                data=data_excel, 
                file_name="sunat_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
