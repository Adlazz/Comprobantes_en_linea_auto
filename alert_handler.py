from selenium.common.exceptions import TimeoutException, NoAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time
import logging

# Configuración del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('alert_handler.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

class AlertHandler:
    """
    Clase para manejar diferentes tipos de alertas y confirmaciones en la aplicación AFIP.
    Implementa múltiples estrategias para asegurar el manejo correcto de las alertas.
    """
    
    def __init__(self, wait_time=10):
        self.wait_time = wait_time
    
    def handle_confirmation(self, driver):
        """
        Maneja las ventanas de confirmación usando múltiples estrategias.
        
        Args:
            driver: WebDriver de Selenium
            
        Returns:
            bool: True si la alerta fue manejada exitosamente, False en caso contrario
        """
        logger.info("Iniciando manejo de confirmación...")
        
        try:
            # Pausa breve para asegurar que la alerta esté presente
            time.sleep(2)
            
            # Lista de estrategias a intentar
            strategies = [
                self._handle_simple_alert,
                self._handle_wait_alert,
                self._handle_popup_window,
                self._handle_js_alert
            ]
            
            # Intentar cada estrategia
            for strategy in strategies:
                try:
                    if strategy(driver):
                        logger.info(f"Confirmación manejada exitosamente usando {strategy.__name__}")
                        return True
                except Exception as e:
                    logger.debug(f"Estrategia {strategy.__name__} falló: {str(e)}")
                    continue
            
            logger.warning("Todas las estrategias de manejo de confirmación fallaron")
            return False
            
        except Exception as e:
            logger.error(f"Error general en manejo de confirmación: {str(e)}")
            return False
    
    def _handle_simple_alert(self, driver):
        """Maneja una alerta simple usando switch_to.alert"""
        try:
            alert = driver.switch_to.alert
            alert.accept()
            return True
        except NoAlertPresentException:
            return False
    
    def _handle_wait_alert(self, driver):
        """Maneja una alerta usando wait explícito"""
        try:
            wait = WebDriverWait(driver, self.wait_time)
            alert = wait.until(EC.alert_is_present())
            alert.accept()
            return True
        except TimeoutException:
            return False
    
    def _handle_popup_window(self, driver):
        """Maneja una ventana emergente cambiando entre ventanas"""
        try:
            main_window = driver.current_window_handle
            WebDriverWait(driver, self.wait_time).until(lambda d: len(d.window_handles) > 1)
            
            for handle in driver.window_handles:
                if handle != main_window:
                    driver.switch_to.window(handle)
                    button = WebDriverWait(driver, self.wait_time).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceptar')]"))
                    )
                    button.click()
                    driver.switch_to.window(main_window)
                    return True
            return False
        except Exception:
            return False
    
    def _handle_js_alert(self, driver):
        """Maneja una alerta usando JavaScript"""
        try:
            driver.execute_script("document.querySelector('button.aceptar').click();")
            return True
        except Exception:
            return False
    
    def handle_error_alert(self, driver):
        """
        Maneja alertas de error específicas.
        
        Args:
            driver: WebDriver de Selenium
            
        Returns:
            tuple: (bool, str) - Éxito del manejo y mensaje de error si existe
        """
        try:
            alert = WebDriverWait(driver, self.wait_time).until(EC.alert_is_present())
            alert_text = alert.text
            alert.accept()
            logger.warning(f"Alerta de error detectada: {alert_text}")
            return True, alert_text
        except TimeoutException:
            return False, None
        except Exception as e:
            logger.error(f"Error manejando alerta de error: {str(e)}")
            return False, str(e)