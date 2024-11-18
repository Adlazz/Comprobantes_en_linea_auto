from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import ElementNotInteractableException, StaleElementReferenceException
import logging
import time

# Configuración del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('element_handler.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

class ElementHandler:
    """
    Clase para manejar interacciones seguras con elementos web en la aplicación AFIP.
    Proporciona métodos para clicks, selecciones y entradas de datos con fallbacks.
    """
    
    def __init__(self, driver, wait):
        """
        Inicializa el manejador de elementos.
        
        Args:
            driver: WebDriver de Selenium
            wait: WebDriverWait configurado
        """
        self.driver = driver
        self.wait = wait
        self.retry_attempts = 3
        self.retry_delay = 1
    
    def safe_click(self, locator, js_fallback=None, description=""):
        """
        Intenta hacer click en un elemento de forma segura.
        
        Args:
            locator: tupla (By.X, "valor") para localizar el elemento
            js_fallback: código JavaScript alternativo para el click
            description: descripción del elemento para logging
            
        Returns:
            bool: True si el click fue exitoso, False si falló
        """
        logger.info(f"Intentando click en elemento: {description}")
        
        for attempt in range(self.retry_attempts):
            try:
                element = self.wait.until(EC.element_to_be_clickable(locator))
                element.click()
                logger.info(f"Click exitoso en {description}")
                return True
            except Exception as e:
                logger.warning(f"Intento {attempt + 1} fallido para click en {description}: {str(e)}")
                if js_fallback:
                    try:
                        self.driver.execute_script(js_fallback)
                        logger.info(f"Click exitoso usando JavaScript en {description}")
                        return True
                    except Exception as js_e:
                        logger.error(f"Error en JavaScript fallback: {str(js_e)}")
                
                if attempt < self.retry_attempts - 1:
                    time.sleep(self.retry_delay)
                    continue
                
                logger.error(f"Todos los intentos de click fallaron para {description}")
                return False
    
    def safe_select(self, element_id, value, description="", scroll_into_view=True):
        """
        Realiza una selección segura en un elemento select.
        
        Args:
            element_id: ID del elemento select
            value: valor a seleccionar
            description: descripción del elemento para logging
            scroll_into_view: si debe desplazarse hasta el elemento
            
        Returns:
            bool: True si la selección fue exitosa, False si falló
        """
        logger.info(f"Intentando seleccionar '{value}' en {description}")
        
        try:
            select_element = self.wait.until(EC.presence_of_element_located((By.ID, element_id)))
            
            if scroll_into_view:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
                time.sleep(0.5)
            
            select = Select(select_element)
            select.select_by_value(value)
            logger.info(f"Selección exitosa de '{value}' en {description}")
            return True
            
        except Exception as e:
            logger.warning(f"Error en selección directa: {str(e)}, intentando JavaScript")
            try:
                js_script = f"""
                    var select = document.getElementById('{element_id}');
                    select.value = '{value}';
                    var event = new Event('change');
                    select.dispatchEvent(event);
                """
                self.driver.execute_script(js_script)
                logger.info(f"Selección exitosa usando JavaScript en {description}")
                return True
            except Exception as js_e:
                logger.error(f"Error en selección por JavaScript: {str(js_e)}")
                return False
    
    def safe_input(self, element_id, value, description=""):
        """
        Ingresa texto de forma segura en un campo de entrada.
        
        Args:
            element_id: ID del elemento input
            value: valor a ingresar
            description: descripción del elemento para logging
            
        Returns:
            bool: True si la entrada fue exitosa, False si falló
        """
        logger.info(f"Intentando ingresar valor en {description}")
        
        try:
            input_element = self.wait.until(EC.presence_of_element_located((By.ID, element_id)))
            input_element.clear()
            input_element.send_keys(str(value).strip())
            logger.info(f"Ingreso exitoso en {description}")
            return True
        except Exception as e:
            logger.error(f"Error en ingreso de datos: {str(e)}")
            return False
    
    def wait_and_find_element(self, locator, timeout=None):
        """
        Espera y encuentra un elemento de forma segura.
        
        Args:
            locator: tupla (By.X, "valor") para localizar el elemento
            timeout: tiempo máximo de espera (opcional)
            
        Returns:
            WebElement o None
        """
        try:
            if timeout:
                element = WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located(locator)
                )
            else:
                element = self.wait.until(EC.presence_of_element_located(locator))
            return element
        except Exception as e:
            logger.error(f"Error encontrando elemento {locator}: {str(e)}")
            return None
    
    def element_exists(self, locator):
        """
        Verifica si un elemento existe en la página.
        
        Args:
            locator: tupla (By.X, "valor") para localizar el elemento
            
        Returns:
            bool: True si existe, False si no
        """
        try:
            self.driver.find_element(*locator)
            return True
        except:
            return False
    
    def wait_for_text(self, locator, text, timeout=10):
        """
        Espera hasta que un elemento contenga un texto específico.
        
        Args:
            locator: tupla (By.X, "valor") para localizar el elemento
            text: texto a esperar
            timeout: tiempo máximo de espera
            
        Returns:
            bool: True si se encontró el texto, False si no
        """
        try:
            WebDriverWait(self.driver, timeout).until(
                lambda driver: text in driver.find_element(*locator).text
            )
            return True
        except:
            return False
    
    def safe_clear_and_send_keys(self, element_id, value, description=""):
        """
        Limpia y envía texto a un elemento de forma segura.
        
        Args:
            element_id: ID del elemento
            value: valor a ingresar
            description: descripción del elemento para logging
            
        Returns:
            bool: True si fue exitoso, False si falló
        """
        logger.info(f"Intentando limpiar e ingresar valor en {description}")
        
        try:
            element = self.wait.until(EC.presence_of_element_located((By.ID, element_id)))
            
            # Intentar limpiar de múltiples formas
            try:
                element.clear()
            except:
                self.driver.execute_script(f"document.getElementById('{element_id}').value = '';")
            
            # Intentar enviar las teclas
            element.send_keys(str(value).strip())
            
            # Verificar que el valor se estableció correctamente
            actual_value = element.get_attribute('value')
            if actual_value == str(value).strip():
                logger.info(f"Ingreso exitoso en {description}")
                return True
            else:
                logger.warning(f"El valor ingresado no coincide en {description}")
                return False
                
        except Exception as e:
            logger.error(f"Error en ingreso de datos: {str(e)}")
            return False
