from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoAlertPresentException
from pathlib import Path
import time
import random
import keyboard  # Para detectar la tecla Esc
from excel_handler import ExcelHandler, FacturaData 
from invoice_processor import InvoiceProcessor
from element_handler import ElementHandler  # Si no lo tienes ya importado
from alert_handler import AlertHandler 

def get_random_user_agent():
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15'
    ]
    return random.choice(user_agents)

def manejar_ventana_confirmacion(driver):
    try:
        # Esperar a que se complete la acción anterior
        time.sleep(2)  # Pequeña pausa para asegurar que la alerta aparezca
        
        # Método 1: Usando alert simple
        try:
            alert = driver.switch_to.alert
            alert.accept()
            print("Alerta aceptada correctamente")
            return True
        except NoAlertPresentException:
            print("No se encontró alerta simple, probando otros métodos...")

        # Método 2: Usando wait explícito para alert
        try:
            alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert.accept()
            print("Alerta aceptada con espera explícita")
            return True
        except TimeoutException:
            print("No se encontró alerta con espera explícita, probando otros métodos...")

        # Método 3: Cambiar a ventana emergente
        try:
            # Obtener el handle de la ventana principal
            main_window = driver.current_window_handle
            
            # Esperar y cambiar a la nueva ventana
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
            
            # Cambiar a la nueva ventana
            for handle in driver.window_handles:
                if handle != main_window:
                    driver.switch_to.window(handle)
                    
                    # Buscar el botón de aceptar en la nueva ventana
                    aceptar_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceptar')]"))
                    )
                    aceptar_btn.click()
                    
                    # Volver a la ventana principal
                    driver.switch_to.window(main_window)
                    print("Ventana emergente manejada correctamente")
                    return True
        except Exception as e:
            print(f"Error al manejar ventana emergente: {str(e)}")

        # Método 4: Intento con JavaScript
        try:
            driver.execute_script("document.querySelector('button.aceptar').click();")
            print("Ventana manejada con JavaScript")
            return True
        except Exception as e:
            print(f"Error al intentar con JavaScript: {str(e)}")

        return False

    except Exception as e:
        print(f"Error general en manejo de ventana: {str(e)}")
        return False

def login_afip(cuit, password):
    options = webdriver.ChromeOptions()
    user_agent = get_random_user_agent()
    options.add_argument(f'user-agent={user_agent}')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": user_agent})
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    time.sleep(random.uniform(1, 3))
    
    try:
        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
        wait = WebDriverWait(driver, 10)

        # Ingresa el CUIT
        text_box = wait.until(EC.presence_of_element_located((By.ID, "F1:username")))
        text_box.clear()
        text_box.send_keys(cuit)
        print(f"CUIT ingresado: {cuit}")

        # Hace clic en el botón "Siguiente"
        siguiente_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnSiguiente")))
        siguiente_button.click()
        print("Se hizo clic en el botón Siguiente")

        # Ingresa la contraseña
        password_field = wait.until(EC.presence_of_element_located((By.ID, "F1:password")))
        password_field.send_keys(password)
        print("Contraseña ingresada")

        # Hace clic en el botón "Ingresar"
        ingresar_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnIngresar")))
        ingresar_button.click()
        print("Se hizo clic en el botón Ingresar")

        time.sleep(5)

        # Hace clic en "Ver Todos" en la página principal
        mis_comprobantes = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Ver todos")))
        mis_comprobantes.click()
        print("Se encontró y abrió 'Ver todos'")

        # Hacer clic en "Comprobantes en Línea"
        comprobantes_link = wait.until(EC.presence_of_element_located((By.XPATH, "//h3[contains(text(),'COMPROBANTES EN LÍNEA')]")))
        
        # Scroll hasta el elemento
        driver.execute_script("arguments[0].scrollIntoView(true);", comprobantes_link)
        time.sleep(1)
      
        # Click en Comprobantes en Línea
        try:
            comprobantes_link.click()
            print("Se hizo clic en Comprobantes en Línea")
        except Exception as e:
            driver.execute_script("arguments[0].click();", comprobantes_link)
            print("Se hizo clic en Comprobantes en Línea usando JavaScript")

        time.sleep(5)

        # Obtener todas las pestañas abiertas
        handles = driver.window_handles
        
        # Cambiar a la última pestaña abierta
        driver.switch_to.window(handles[-1])
        print("Cambiado a la nueva pestaña")
        
        try:
            # Esperar y hacer clic en el botón de la empresa
            wait = WebDriverWait(driver, 10)
            empresa_button = wait.until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, "input.btn_empresa[value='LAZZARINI&LAZZARINI S.R.L.']"
            )))
            
            driver.execute_script("arguments[0].scrollIntoView(true);", empresa_button)
            time.sleep(1)
            
            empresa_button.click()
            print("Se hizo clic en el botón de la empresa")
            
        except Exception as e:
            print(f"Error al hacer clic en el botón de la empresa: {str(e)}")
            try:
                button = driver.find_element(By.CSS_SELECTOR, "input.btn_empresa[value='LAZZARINI&LAZZARINI S.R.L.']")
                driver.execute_script("arguments[0].click();", button)
                print("Se hizo clic usando JavaScript en el botón de la empresa")
            except Exception as js_e:
                print(f"Error en el clic por JavaScript: {str(js_e)}")

        # Loop principal para procesar facturas
        while True:
            try:
                # Procesar lote de facturas
                facturar(driver)
                
                # Verificar si quedan más facturas
                excel_handler = ExcelHandler("facturador_test.xlsx")
                if excel_handler.load_excel():
                    facturas_pendientes = excel_handler.get_facturas_pendientes()
                    if not facturas_pendientes:
                        print("\nNo quedan facturas pendientes. Proceso completado.")
                        break
                    else:
                        print(f"\nQuedan {len(facturas_pendientes)} facturas pendientes. Continuando...")
                        continue

            except KeyboardInterrupt:
                print("\nProceso interrumpido por el usuario.")
                break
            except Exception as e:
                print(f"\nError en el loop principal: {str(e)}")
                driver.save_screenshot("error_loop_principal.png")
                break

        # Esperar la tecla Esc para cerrar
        print("\nProceso completado. Mantén la sesión abierta, presiona 'Esc' para cerrar.")
        while True:
            if keyboard.is_pressed('esc'):
                print("Tecla 'Esc' presionada. Cerrando sesión.")
                break

    except Exception as e:
        driver.save_screenshot(f"afip_login_error_{cuit}.png")
        print(f"Error durante el login: {str(e)}. Se guardó una captura de pantalla.")
        return None

    finally:
        driver.quit()
        print("Sesión cerrada.")

def facturar(driver):
    try:
        # Primero definir wait
        wait = WebDriverWait(driver, 10)

        # Luego inicializar handlers con wait ya definido
        element_handler = ElementHandler(driver, wait)
        alert_handler = AlertHandler()
        invoice_processor = InvoiceProcessor(driver, element_handler, alert_handler)

        # Instanciar el ExcelHandler
        excel_handler = ExcelHandler("facturador_test.xlsx")
        if not excel_handler.load_excel():
            raise Exception("No se pudo cargar el archivo Excel")

        # Obtener facturas pendientes
        facturas_pendientes = excel_handler.get_facturas_pendientes()
        if not facturas_pendientes:
            print("No hay facturas pendientes para procesar")
            return

        # Por cada factura pendiente
        for factura in facturas_pendientes:
            try:
                # Validar datos de la factura
                if not excel_handler.validate_factura_data(factura):
                    print(f"Factura para {factura.cliente} no pasó la validación, continuando con la siguiente...")
                    continue

                # Hace clic en "Generar Comprobantes"
                generar_comprobantes = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Generar Comprobantes")))
                time.sleep(1)
                
                try:
                    generar_comprobantes.click()
                except:
                    driver.execute_script("arguments[0].click();", generar_comprobantes)
                print("Se encontró y abrió 'Generar Comprobantes'")
                
                # Seleccionar Punto de Venta
                select_element = wait.until(EC.presence_of_element_located((By.ID, "puntodeventa")))
                select = Select(select_element)
                time.sleep(2)
                
                try:
                    select.select_by_value("4")
                    print("Se seleccionó el Punto de Venta 4")
                except Exception as e:
                    print(f"Error al seleccionar Punto de Venta: {str(e)}")
                    driver.execute_script("document.getElementById('puntodeventa').value='4';")
                    driver.execute_script("obtenerDenominacion(formulario,'puntodeventa');")

                # Seleccionar tipo de comprobante según condición IVA
                wait.until(EC.visibility_of_element_located((By.ID, "universocomprobante")))
                time.sleep(2)

                try:
                    # Determinar el valor según la condición IVA
                    if factura.cond_iva.upper() in ['RI', 'M']:
                        valor_comprobante = "10"  # Factura A
                        print(f"Cliente {factura.cliente} es {factura.cond_iva} - Seleccionando Factura A")
                    else:
                        valor_comprobante = "19"  # Factura B 
                        print(f"Cliente {factura.cliente} es {factura.cond_iva} - Seleccionando Factura B")

                    # Intentar selección directa
                    select_comprobante = Select(driver.find_element(By.ID, "universocomprobante"))
                    select_comprobante.select_by_value(valor_comprobante)
                    time.sleep(1)
                
                except Exception as e:
                    print(f"Error al seleccionar tipo de comprobante: {str(e)}")
                    # Plan B: JavaScript
                    script = f"""
                        var select = document.getElementById("universocomprobante");
                        select.value = "{valor_comprobante}";
                        var event = new Event('change');
                        select.dispatchEvent(event);
                        actualizarDescripcionTC(select.selectedIndex);
                    """
                    driver.execute_script(script)
                    time.sleep(1)
                    
                # Hacer clic en el botón Continuar
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    continuar_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad
                    
                    # Intentar clic normal
                    try:
                        continuar_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("validarCampos();")
                    
                    print("Se hizo clic en Continuar")
                    time.sleep(2)  # Esperar a que cargue la siguiente página
                    
                except Exception as e:
                    print(f"Error al hacer clic en Continuar: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_continuar_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior

                # Ingresa Fecha
                text_box = wait.until(EC.presence_of_element_located((By.ID, "fc")))
                text_box.clear()
                fecha_formateada = factura.fecha.strftime("%d/%m/%Y")  # Convertir a formato dd/mm/yyyy
                text_box.send_keys(fecha_formateada)
                print(f"Fecha ingresada: {fecha_formateada}")

                # Seleccionar concepto Servicio
                wait.until(EC.visibility_of_element_located((By.ID, "idconcepto")))
                time.sleep(2)

                try:
                    # Intentar selección directa
                    select_concepto = Select(driver.find_element(By.ID, "idconcepto"))
                    select_concepto.select_by_value("2")
                    time.sleep(1)
                    print("Se seleccionó el Concepto Servicio")
                
                except Exception as e:
                    print(f"Error al seleccionar Concepto: {str(e)}")
                    # Plan B: JavaScript
                    script = """
                        var select = document.getElementById("idconcepto");
                        select.value = "2";
                        var event = new Event('change');
                        select.dispatchEvent(event);
                        mostrarOcultar(select.value);
                    """
                    driver.execute_script(script)
                    time.sleep(1)

                # Ingresa Fecha Desde
                text_box = wait.until(EC.presence_of_element_located((By.ID, "fsd")))
                text_box.clear()
                fecha_formateada = factura.fecha.strftime("%d/%m/%Y")  # Convertir a formato dd/mm/yyyy
                text_box.send_keys(fecha_formateada)
                print(f"Fecha Desde ingresada: {fecha_formateada}")


                # Ingresa Fecha Hasta
                text_box = wait.until(EC.presence_of_element_located((By.ID, "fsh")))
                text_box.clear()
                fecha_formateada = factura.fecha.strftime("%d/%m/%Y")  # Convertir a formato dd/mm/yyyy
                text_box.send_keys(fecha_formateada)
                print(f"Fecha Hasta ingresada: {fecha_formateada}")


                # Ingresa Fecha Vto. Pago
                text_box = wait.until(EC.presence_of_element_located((By.ID, "vencimientopago")))
                text_box.clear()
                fecha_formateada = factura.fecha.strftime("%d/%m/%Y")  # Convertir a formato dd/mm/yyyy
                text_box.send_keys(fecha_formateada)
                print(f"Fecha Vto. Pago ingresada: {fecha_formateada}")

                # Seleccionar Actividad
                select_element = wait.until(EC.presence_of_element_located((By.ID, "actiAsociadaId")))
                select_concepto = Select(select_element)
                time.sleep(2)
                
                try:
                    select_concepto.select_by_value("682091")
                    print("Se seleccionó la Actividad")
                except Exception as e:
                    print(f"Error al seleccionar Concepto: {str(e)}")
                    driver.execute_script("document.getElementById('actiAsociadaId').value='682091';")
                    driver.execute_script("obtenerDenominacion(formulario,'actiAsociadaId');")

                # Hacer clic en el botón Continuar
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    continuar_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad
                    
                    # Intentar clic normal
                    try:
                        continuar_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("validarCampos();")
                    
                    print("Se hizo clic en Continuar")
                    time.sleep(2)  # Esperar a que cargue la siguiente página
                   
                except Exception as e:
                    print(f"Error al hacer clic en Continuar: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_continuar_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior

                # Seleccionar condición IVA                 
                wait.until(EC.visibility_of_element_located((By.ID, "idivareceptor")))                 
                time.sleep(2)                  

                try:
                    # Determinar el valor según la condición IVA
                    if factura.cond_iva.upper() == 'RI':
                        condicion_iva = "1"  # IVA Responsable Inscripto
                        print(f"Cliente {factura.cliente} es Responsable Inscripto")
                    elif factura.cond_iva.upper() == 'M':
                        condicion_iva = "6"  # Responsable Monotributo 
                        print(f"Cliente {factura.cliente} es Monotributista")
                    elif factura.cond_iva.upper() == 'CF':
                        condicion_iva = "5"  # Consumidor Final
                        print(f"Cliente {factura.cliente} es Consumidor Final")
                    elif factura.cond_iva.upper() == 'E':
                        condicion_iva = "4"  # IVA Sujeto Exento
                        print(f"Cliente {factura.cliente} es IVA Exento")

                    # Intentar selección directa
                    select_comprobante = Select(driver.find_element(By.ID, "idivareceptor"))
                    select_comprobante.select_by_value(condicion_iva)
                    time.sleep(1)
                            
                except Exception as e:
                    print(f"Error al seleccionar condición IVA: {str(e)}")
                    # Plan B: JavaScript
                    script = f"""
                        var select = document.getElementById("idivareceptor");
                        select.value = "{condicion_iva}";
                        var event = new Event('change');
                        select.dispatchEvent(event);
                    """
                    driver.execute_script(script)
                    time.sleep(1)

                # Ingresa el CUIT
                text_box = wait.until(EC.presence_of_element_located((By.ID, "nrodocreceptor")))
                text_box.clear()
                text_box.send_keys(factura.cuit)
                print(f"CUIT ingresado: {factura.cuit}")
                time.sleep(8)

                # Hacer click en checkbox Contado
                checkbox = wait.until(EC.element_to_be_clickable((By.ID, "formadepago1")))
                time.sleep(1)
                checkbox.click()
                print("Se hizo clic en Contado como Condición de Venta")

                # Hacer clic en el botón Continuar
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    continuar_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad
                    
                    # Intentar clic normal
                    try:
                        continuar_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("validarCampos();")
                    
                    print("Se hizo clic en Continuar")
                    time.sleep(2)  # Esperar a que cargue la siguiente página
                   
                except Exception as e:
                    print(f"Error al hacer clic en Continuar: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_continuar_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior

                # Escribir concepto
                text_box = wait.until(EC.presence_of_element_located((By.ID, "detalle_descripcion1")))
                text_box.clear()
                text_box.send_keys(f"Comisiones por cobranzas Mes de {factura.periodo} - Rendición N° {str(factura.rendicion).strip()}")
                print(f"Concepto ingresado correctamente")

                # Seleccionar Unidad de Medida
                select_element = wait.until(EC.presence_of_element_located((By.ID, "detalle_medida1")))
                select_um = Select(select_element)
                time.sleep(2)
                
                try:
                    select_um.select_by_value("98")
                    print("Se seleccionó la Unidad de Medida")
                except Exception as e:
                    print(f"Error al seleccionar Unidad de Medida: {str(e)}")
                    driver.execute_script("document.getElementById('detalle_medida1').value='98';")
                    driver.execute_script("obtenerDenominacion(formulario,'detalle_medida1');")

                # Escribir importe
                text_box = wait.until(EC.presence_of_element_located((By.ID, "detalle_precio1")))
                text_box.clear()
                text_box.send_keys(factura.importe)
                print(f"Importe ingresado correctamente")
                time.sleep(2)

                # NUEVO CÓDIGO: Seleccionar IVA 21% solo si es factura B
                if factura.cond_iva.upper() not in ['RI', 'M']:  # Si no es RI ni M, es factura B
                    try:
                        # Esperar a que el elemento esté presente
                        select_iva = wait.until(EC.presence_of_element_located((By.ID, "detalle_tipo_iva1")))
                        select_iva_dropdown = Select(select_iva)
                        time.sleep(1)
                        
                        # Intentar selección directa del 21%
                        try:
                            select_iva_dropdown.select_by_value("5")  # 5 corresponde al 21%
                            print("Se seleccionó IVA 21%")
                        except Exception as e:
                            print(f"Error al seleccionar IVA por método directo: {str(e)}")
                            # Plan B: JavaScript
                            script = """
                                var select = document.getElementById("detalle_tipo_iva1");
                                select.value = "5";
                                var event = new Event('change');
                                select.dispatchEvent(event);
                                calcularSubtotalDetalle(1);
                            """
                            driver.execute_script(script)
                            print("Se seleccionó IVA 21% usando JavaScript")
                        
                        time.sleep(2)  # Esperar a que se actualicen los cálculos
                        
                    except Exception as e:
                        print(f"Error al seleccionar porcentaje de IVA: {str(e)}")
                        driver.save_screenshot(f"error_seleccion_iva_{factura.cliente}.png")

                # Hacer clic en el botón Continuar
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    continuar_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad
                    
                    # Intentar clic normal
                    try:
                        continuar_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("validarCampos();")
                    
                    print("Se hizo clic en Continuar")
                    time.sleep(2)  # Esperar a que cargue la siguiente página
                   
                except Exception as e:
                    print(f"Error al hacer clic en Continuar: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_continuar_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior

                # Hacer clic en el botón Confirmar Datos
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    confirmar_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Confirmar Datos...' and @onclick='confirmar();']"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad
                    
                    # Intentar clic normal
                    try:
                        confirmar_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("confirmar();")
                    
                    print("Se hizo clic en Confirmar Datos")
                    time.sleep(2)  # Esperar a que aparezca la ventana de confirmación
                   
                except Exception as e:
                    print(f"Error al hacer clic en Confirmar Datos: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_confirmar_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior  

                # Manejar la ventana de confirmación
                if manejar_ventana_confirmacion(driver):
                    print("Ventana de confirmación manejada exitosamente")
                    time.sleep(5)  # Esperar a que se procese la confirmación
                else:
                    print("No se pudo manejar la ventana de confirmación")
                    driver.save_screenshot(f"error_ventana_confirmacion_{factura.cliente}.png")
                    raise Exception("Error al manejar la ventana de confirmación")
                
                # Intentar hacer click en el botón Imprimir
                try:
                    print("Buscando botón Imprimir...")
                    # Esperar a que la página se actualice
                    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
                    
                    # Intentar diferentes localizadores para el botón imprimir
                    locators = [
                        "//input[@type='button' and @value='Imprimir...']",
                        "//*[@id='botones_comprobante']/input",
                        "//input[contains(@onclick, 'imprimirComprobante.do')]",
                        "//input[@type='button'][contains(@value, 'Imprimir')]"
                    ]

                    imprimir_btn = None
                    for xpath in locators:
                        try:
                            imprimir_btn = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, xpath))
                            )
                            if imprimir_btn.is_displayed():
                                print(f"Botón Imprimir encontrado con localizador: {xpath}")
                                break
                        except:
                            continue

                    if imprimir_btn and imprimir_btn.is_displayed():
                        # Scroll al botón
                        driver.execute_script("arguments[0].scrollIntoView(true);", imprimir_btn)
                        time.sleep(1)

                        try:
                            print("Intentando click en botón Imprimir...")
                            imprimir_btn.click()
                        except:
                            print("Click directo falló, intentando con JavaScript...")
                            onclick = imprimir_btn.get_attribute('onclick')
                            if onclick:
                                driver.execute_script(onclick)
                            else:
                                driver.execute_script("arguments[0].click();", imprimir_btn)

                        print("Click en Imprimir realizado")
                        time.sleep(5)  # Esperar a que se genere el PDF
                              
                        # Usar los métodos del InvoiceProcessor para manejar el PDF
                        downloads_folder = str(Path.home() / "Downloads")
                        destino_folder = str(Path.home() / "Desktop")

                        print("Verificando descarga del PDF...")
                        max_intentos_descarga = 3
                        for intento in range(max_intentos_descarga):
                            try:
                                if invoice_processor._verify_download_started(downloads_folder):
                                    print("Descarga detectada, esperando que complete...")
                                    time.sleep(5)  # Esperar a que se complete la descarga
                                    
                                    print("Intentando procesar el archivo...")
                                    if invoice_processor._process_downloaded_file(downloads_folder, destino_folder, factura):
                                        print("Archivo PDF procesado exitosamente")
                                        break
                                    else:
                                        print("Error procesando el archivo PDF")
                                        if intento == max_intentos_descarga - 1:
                                            raise Exception("No se pudo procesar el PDF después de varios intentos")
                                else:
                                    print(f"Intento {intento + 1}: No se detectó la descarga aún")
                                    if intento == max_intentos_descarga - 1:
                                        raise Exception("No se detectó la descarga del PDF después de varios intentos")
                                    time.sleep(2)
                            except Exception as e:
                                print(f"Error en intento {intento + 1} de procesar PDF: {str(e)}")
                                if intento == max_intentos_descarga - 1:
                                    raise

                        # Solo si todo el proceso fue exitoso, hacer click en Menú Principal
                        print("Procediendo a Menú Principal...")

                    else:
                        print("No se pudo encontrar el botón Imprimir")
                        driver.save_screenshot("error_no_boton_imprimir.png")

                except Exception as e:
                    print(f"Error al intentar imprimir: {str(e)}")
                    driver.save_screenshot("error_imprimir.png")
                
                # Hacer clic en el botón Menú Principal
                try:
                    # Esperar a que el botón esté presente y sea clickeable
                    menu_principal_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//input[@type='button' and @value='Menú Principal' and contains(@onclick, 'menu_ppal.jsp')]"
                    )))
                    time.sleep(1)  # Pequeña pausa para estabilidad

                    # Opción 1: Usar JavaScript para scroll hasta el elemento
                    driver.execute_script("arguments[0].scrollIntoView(true);", menu_principal_button)
                    # Pequeña pausa para que termine la animación del scroll
                    time.sleep(1)
                    
                    # Intentar clic normal
                    try:
                        menu_principal_button.click()
                    except:
                        # Si falla, intentar con JavaScript
                        driver.execute_script("parent.location.href='menu_ppal.jsp'")
                    
                    print("Se hizo clic en Menú Principal")
                    time.sleep(4)  # Esperar a que cargue la siguiente página         
                
                except Exception as e:
                    print(f"Error al hacer clic en Menú Principal: {str(e)}")
                    # Capturar screenshot si hay error
                    driver.save_screenshot(f"error_menu_principal_{factura.cliente}.png")
                    raise  # Re-lanzar el error para que se maneje en el bloque try/except superior  

                # Si todo sale bien, marcar como realizada
                if excel_handler.marcar_como_realizada(factura):
                    print(f"Factura para {factura.cliente} procesada exitosamente")
                    time.sleep(2)  # Esperar antes de la siguiente factura
              
            except Exception as e:
                print(f"Error procesando factura para {factura.cliente}: {str(e)}")
                driver.save_screenshot(f"error_factura_{factura.cliente}.png")
                continue
        
    except Exception as e:
        print(f"Error en la función facturar: {str(e)}")
        driver.save_screenshot("error_facturar.png")

# Ejemplo de uso
if __name__ == "__main__":
    cuit = "20146327401"
    password = "Ramolaz2024"
    login_afip(cuit, password)
