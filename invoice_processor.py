from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
import logging
import time
import os
import time
import shutil
from pathlib import Path

# Configuración del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('invoice_processor.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

class InvoiceProcessor:
    """
    Clase para procesar facturas en el sistema AFIP.
    Maneja todo el flujo de creación de facturas desde el inicio hasta el fin.
    """
    
    def __init__(self, driver, element_handler, alert_handler):
        """
        Inicializa el procesador de facturas.
        
        Args:
            driver: WebDriver de Selenium
            element_handler: Instancia de ElementHandler
            alert_handler: Instancia de AlertHandler
        """
        self.driver = driver
        self.handler = element_handler
        self.alert_handler = alert_handler
        self.wait = WebDriverWait(driver, 10)

    def process_invoice(self, factura):
        """
        Procesa una factura completa.
        
        Args:
            factura: Objeto FacturaData con la información de la factura
            
        Returns:
            bool: True si el proceso fue exitoso, False si falló
        """
        logger.info(f"Iniciando procesamiento de factura para {factura.cliente}")
        
        try:
            # Secuencia de pasos para procesar la factura
            steps = [
                self._init_invoice,
                self._fill_basic_info,
                self._fill_dates,
                self._fill_client_info,
                self._fill_invoice_details,
                self._confirm_invoice
            ]
            
            # Ejecutar cada paso
            for step in steps:
                if not step(factura):
                    logger.error(f"Falló el paso {step.__name__} para {factura.cliente}")
                    self.driver.save_screenshot(f"error_{step.__name__}_{factura.cliente}.png")
                    return False
                    
            logger.info(f"Factura procesada exitosamente para {factura.cliente}")
            return True
            
        except Exception as e:
            logger.error(f"Error procesando factura para {factura.cliente}: {str(e)}")
            self.driver.save_screenshot(f"error_factura_{factura.cliente}.png")
            return False

    def _init_invoice(self, factura):
        """Inicia el proceso de facturación"""
        try:
            # Click en Generar Comprobantes
            if not self.handler.safe_click(
                (By.LINK_TEXT, "Generar Comprobantes"),
                description="Botón Generar Comprobantes"
            ):
                return False

            # Seleccionar Punto de Venta
            if not self.handler.safe_select(
                "puntodeventa",
                "4",
                description="Punto de Venta"
            ):
                return False

            # Seleccionar tipo de comprobante
            tipo_comprobante = "10" if factura.cond_iva.upper() in ['RI', 'M'] else "19"
            if not self.handler.safe_select(
                "universocomprobante",
                tipo_comprobante,
                description="Tipo de Comprobante"
            ):
                return False

            # Click en Continuar
            return self.handler.safe_click(
                (By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"),
                js_fallback="validarCampos();",
                description="Botón Continuar"
            )

        except Exception as e:
            logger.error(f"Error en inicialización de factura: {str(e)}")
            return False

    def _fill_basic_info(self, factura):
        """Completa la información básica de la factura"""
        try:
            # Seleccionar concepto
            if not self.handler.safe_select(
                "idconcepto",
                "2",  # Servicio
                description="Concepto"
            ):
                return False

            # Seleccionar Actividad
            if not self.handler.safe_select(
                "actiAsociadaId",
                "682091",
                description="Actividad"
            ):
                return False

            return True

        except Exception as e:
            logger.error(f"Error en información básica: {str(e)}")
            return False

    def _fill_dates(self, factura):
        """Completa las fechas de la factura"""
        try:
            fecha_formateada = factura.fecha.strftime("%d/%m/%Y")
            
            # Lista de campos de fecha
            date_fields = [
                ("fc", "Fecha de Comprobante"),
                ("fsd", "Fecha Desde"),
                ("fsh", "Fecha Hasta"),
                ("vencimientopago", "Fecha Vencimiento")
            ]
            
            # Llenar todos los campos de fecha
            for field_id, description in date_fields:
                if not self.handler.safe_input(
                    field_id,
                    fecha_formateada,
                    description=description
                ):
                    return False

            return True

        except Exception as e:
            logger.error(f"Error en fechas: {str(e)}")
            return False

    def _fill_client_info(self, factura):
        """Completa la información del cliente"""
        try:
            # Primero hacer clic en Continuar
            if not self.handler.safe_click(
                (By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"),
                js_fallback="validarCampos();",
                description="Botón Continuar antes de Condición IVA"
            ):
                logger.error("Error al hacer clic en Continuar antes de Condición IVA")
                return False

            # Agregar una pausa para asegurar que la página se cargue
            time.sleep(2)
                
            # Seleccionar condición IVA
            condicion_iva_map = {
                'RI': "1",  # Responsable Inscripto
                'M': "6",   # Monotributista
                'CF': "5",  # Consumidor Final
                'E': "4"    # Exento
            }
            
            if not self.handler.safe_select(
                "idivareceptor",
                condicion_iva_map[factura.cond_iva.upper()],
                description="Condición IVA"
            ):
                return False

            # Ingresar CUIT
            if not self.handler.safe_input(
                "nrodocreceptor",
                factura.cuit,
                description="CUIT"
            ):
                return False

            # Seleccionar Contado
            if not self.handler.safe_click(
                (By.ID, "formadepago1"),
                description="Checkbox Contado"
            ):
                return False

            # Hacer clic en Continuar después de la información del cliente
            if not self.handler.safe_click(
                (By.XPATH, "//input[@type='button' and @value='Continuar >' and @onclick='validarCampos();']"),
                js_fallback="validarCampos();",
                description="Botón Continuar después de info cliente"
            ):
                logger.error("Error al hacer clic en Continuar después de info cliente")
                return False

            # Agregar pausa para asegurar que la página se cargue
            time.sleep(2)

            return True

        except Exception as e:
            logger.error(f"Error en información del cliente: {str(e)}")

    def _fill_invoice_details(self, factura):
        """Completa los detalles de la factura"""
        try:
            # Descripción
            descripcion = f"Comisiones por cobranzas Mes de {factura.periodo} - Rendición N° {str(factura.rendicion).strip()}"
            if not self.handler.safe_input(
                "detalle_descripcion1",
                descripcion,
                description="Descripción"
            ):
                return False

            # Unidad de medida
            if not self.handler.safe_select(
                "detalle_medida1",
                "98",
                description="Unidad de Medida"
            ):
                return False

            # Importe
            if not self.handler.safe_input(
                "detalle_precio1",
                factura.importe,
                description="Importe"
            ):
                return False

            # Si es factura B, seleccionar IVA 21%
            if factura.cond_iva.upper() not in ['RI', 'M']:
                if not self.handler.safe_select(
                    "detalle_tipo_iva1",
                    "5",  # 21%
                    description="Tipo de IVA"
                ):
                    return False

            # Click en botón Continuar después del importe
            if not self.handler.safe_click(
                (By.XPATH, "//input[@type='button' and @value='Continuar >']"),
                description="Botón Continuar después de importe"
            ):
                logger.error("Error al hacer clic en Continuar después del importe")
                return False

            # Pequeña pausa para asegurar que la página se actualice
            time.sleep(2)

            return True

        except Exception as e:
            logger.error(f"Error en detalles de factura: {str(e)}")
            return False

    def _confirm_invoice(self, factura):
        """Confirma y finaliza la factura, maneja el PDF descargado"""
        try:
            # Definir rutas al inicio
            downloads_folder = str(Path.home() / "Downloads")
            destino_folder = str(Path.home() / "Desktop")  # Usar Path para mayor compatibilidad

            # Click en Confirmar Datos
            if not self.handler.safe_click(
                (By.XPATH, "//input[@type='button' and @value='Confirmar Datos...' and @onclick='confirmar();']"),
                js_fallback="confirmar();",
                description="Botón Confirmar Datos"
            ):
                return False

            # Manejar ventana de confirmación
            if not self.alert_handler.handle_confirmation(self.driver):
                logger.error("Error manejando ventana de confirmación")
                return False
                
            logger.info("Esperando a que la página se actualice después de la confirmación...")
                
            # Espera más larga y explícita para la actualización de la página
            try:
                self.wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
                logger.info("Página completamente cargada")
            except:
                logger.warning("No se pudo verificar el estado de carga de la página")

            # Espera base más larga
            time.sleep(5)

            # Verificar el HTML actual
            page_source = self.driver.page_source
            logger.info("Estructura actual de la página:")
            logger.info(page_source)

            # Lista reducida de localizadores más específicos
            locators = [
                (By.XPATH, "//input[@type='button' and @value='Imprimir...']"),  # Más específico
                (By.XPATH, "//*[@id='botones_comprobante']/input[@type='button']"),  # Por ID y tipo
                (By.XPATH, "//input[contains(@onclick, 'imprimirComprobante.do')]")  # Por onclick
            ]

            # Buscar todos los botones en la página
            all_buttons = self.driver.find_elements(By.TAG_NAME, "input")
            logger.info("Botones encontrados en la página:")
            for btn in all_buttons:
                try:
                    btn_info = {
                        'type': btn.get_attribute('type'),
                        'value': btn.get_attribute('value'),
                        'onclick': btn.get_attribute('onclick'),
                        'id': btn.get_attribute('id'),
                        'class': btn.get_attribute('class'),
                        'is_displayed': btn.is_displayed(),
                        'is_enabled': btn.is_enabled()
                    }
                    logger.info(f"Botón encontrado: {btn_info}")
                except:
                    continue

            # Intentar la búsqueda del botón
            max_intentos = 5
            for intento in range(max_intentos):
                try:
                    logger.info(f"\nIntento {intento + 1} de encontrar el botón Imprimir")
                    
                    for locator in locators:
                        try:
                            logger.info(f"Probando localizador: {locator}")
                            
                            # Espera explícita más larga
                            wait = WebDriverWait(self.driver, 10)
                            imprimir_btn = wait.until(
                                EC.presence_of_element_located(locator)
                            )
                            
                            logger.info(f"Botón encontrado con localizador {locator}")
                            logger.info(f"Atributos del botón:")
                            logger.info(f"- Visible: {imprimir_btn.is_displayed()}")
                            logger.info(f"- Habilitado: {imprimir_btn.is_enabled()}")
                            logger.info(f"- HTML: {imprimir_btn.get_attribute('outerHTML')}")
                            
                            if not imprimir_btn.is_displayed():
                                logger.info("Botón no visible, intentando siguiente localizador")
                                continue
                            
                            # Scroll con más margen
                            self.driver.execute_script("""
                                arguments[0].scrollIntoView(true);
                                window.scrollBy(0, -100);
                            """, imprimir_btn)
                            time.sleep(2)
                            
                            # Intentar click directo con espera explícita
                            wait.until(EC.element_to_be_clickable(locator))
                            
                            try:
                                logger.info("Intentando click directo...")
                                imprimir_btn.click()
                            except Exception as click_e:
                                logger.warning(f"Click directo falló: {click_e}")
                                
                                try:
                                    logger.info("Intentando click con JavaScript...")
                                    onclick = imprimir_btn.get_attribute('onclick')
                                    if onclick:
                                        logger.info(f"Ejecutando onclick: {onclick}")
                                        self.driver.execute_script(onclick)
                                    else:
                                        self.driver.execute_script("arguments[0].click();", imprimir_btn)
                                except Exception as js_e:
                                    logger.warning(f"Click JavaScript falló: {js_e}")
                                    
                                    try:
                                        logger.info("Intentando click con Actions...")
                                        ActionChains(self.driver).move_to_element(imprimir_btn).pause(1).click().perform()
                                    except Exception as action_e:
                                        logger.warning(f"Click Actions falló: {action_e}")
                                        continue
                            
                            # Verificar si el click fue exitoso
                            logger.info("Verificando si se inició la descarga...")
                            if self._verify_download_started(downloads_folder):
                                logger.info("¡Descarga iniciada correctamente!")
                                time.sleep(5)
                                
                                # Procesar el archivo descargado
                                if self._process_downloaded_file(downloads_folder, destino_folder, factura):
                                    # Solo si todo el proceso fue exitoso, hacer click en Menú Principal
                                    return self.handler.safe_click(
                                        (By.XPATH, "//input[@value='Menú Principal']"),
                                        js_fallback="parent.location.href='menu_ppal.jsp'",
                                        description="Botón Menú Principal"
                                    )
                            else:
                                logger.warning("No se detectó inicio de descarga")
                                
                        except Exception as e:
                            logger.warning(f"Error con localizador {locator}: {str(e)}")
                            continue
                    
                    time.sleep(2)
                    
                except Exception as e:
                    logger.warning(f"Error en intento {intento + 1}: {str(e)}")
                    if intento == max_intentos - 1:
                        logger.error("No se pudo hacer click en Imprimir")
                        self.driver.save_screenshot("error_imprimir.png")
                        return False

            return False

        except Exception as e:
            logger.error(f"Error en confirmación de factura: {str(e)}")
            self.driver.save_screenshot("error_confirmacion.png")
            return False

    def _verify_download_started(self, downloads_folder):
        """Verifica si la descarga comenzó"""
        try:
            # Obtener lista inicial de archivos
            initial_files = set(os.listdir(downloads_folder))
            
            # Esperar hasta 10 segundos por un nuevo archivo
            for _ in range(10):
                time.sleep(1)
                current_files = set(os.listdir(downloads_folder))
                new_files = current_files - initial_files
                
                # Verificar si hay nuevos archivos PDF o archivos temporales de descarga
                if any(f.endswith('.pdf') or f.endswith('.crdownload') for f in new_files):
                    return True
                    
            return False
        except Exception as e:
            logger.error(f"Error verificando descarga: {str(e)}")
            return False

    def _process_downloaded_file(self, downloads_folder, destino_folder, factura):
        """Procesa el archivo descargado"""
        try:
            # Buscar el archivo PDF más reciente
            archivos = [f for f in os.listdir(downloads_folder) if f.endswith('.pdf')]
            if not archivos:
                raise Exception("No se encontró archivo PDF en Downloads")
                
            archivo_reciente = max([os.path.join(downloads_folder, f) for f in archivos], 
                                key=os.path.getctime)

            # Crear nuevo nombre de archivo
            nuevo_nombre = f"{factura.cliente} x Honorarios {factura.periodo} - Rendición N° {factura.rendicion}.pdf"
            nuevo_nombre = nuevo_nombre.replace('/', '-').replace('\\', '-')
            
            # Asegurar que existe la carpeta destino
            os.makedirs(destino_folder, exist_ok=True)
            
            # Ruta completa del archivo destino
            archivo_destino = os.path.join(destino_folder, nuevo_nombre)

            # Mover y renombrar el archivo
            shutil.move(archivo_reciente, archivo_destino)
            logger.info(f"Archivo movido y renombrado exitosamente: {nuevo_nombre}")
            return True

        except Exception as e:
            logger.error(f"Error procesando archivo descargado: {str(e)}")
            raise