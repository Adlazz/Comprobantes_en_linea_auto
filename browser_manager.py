from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import random
import time
import logging

# Configuración del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('browser_manager.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

class BrowserManager:
    """
    Clase para manejar la configuración y gestión del navegador Chrome.
    Proporciona métodos para inicializar y configurar el navegador de manera segura.
    """
    
    def __init__(self):
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15'
        ]
        self.driver = None
        self.wait = None

    def setup_driver(self, wait_time=10):
        """
        Configura y retorna una instancia de Chrome WebDriver con configuraciones anti-detección.
        
        Args:
            wait_time: Tiempo de espera para WebDriverWait
            
        Returns:
            tuple: (WebDriver, WebDriverWait)
        """
        try:
            logger.info("Iniciando configuración del navegador...")
            
            options = self._configure_chrome_options()
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            
            # Configurar el tamaño de la ventana y maximizar
            driver.set_window_size(1920, 1080)
            driver.maximize_window()
            
            # Configurar wait global
            wait = WebDriverWait(driver, wait_time)
            
            self.driver = driver
            self.wait = wait
            
            logger.info("Navegador configurado exitosamente")
            return driver, wait
            
        except Exception as e:
            logger.error(f"Error configurando el navegador: {str(e)}")
            raise

    def _configure_chrome_options(self):
        """
        Configura las opciones de Chrome para evitar detección de automatización.
        
        Returns:
            ChromeOptions: Opciones configuradas
        """
        options = webdriver.ChromeOptions()
        user_agent = random.choice(self.user_agents)
        
        # Configuraciones básicas
        options.add_argument(f'user-agent={user_agent}')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-popup-blocking')
        options.add_argument('--disable-infobars')
        
        # Configuraciones experimentales
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        
        return options

    def navigate_to(self, url, retry_count=3):
        """
        Navega a una URL de forma segura con reintentos.
        
        Args:
            url: URL a navegar
            retry_count: Número de intentos
            
        Returns:
            bool: True si la navegación fue exitosa
        """
        if not self.driver:
            logger.error("Driver no inicializado")
            return False
            
        for attempt in range(retry_count):
            try:
                logger.info(f"Navegando a {url} (intento {attempt + 1})")
                self.driver.get(url)
                return True
            except Exception as e:
                logger.warning(f"Error en navegación (intento {attempt + 1}): {str(e)}")
                if attempt < retry_count - 1:
                    time.sleep(2)
                    continue
                logger.error(f"Falló la navegación a {url} después de {retry_count} intentos")
                return False

    def wait_for_element(self, locator, timeout=10):
        """
        Espera por un elemento específico.
        
        Args:
            locator: tupla (By.X, "valor")
            timeout: tiempo máximo de espera
            
        Returns:
            WebElement o None
        """
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(locator)
            )
        except Exception as e:
            logger.error(f"Timeout esperando elemento {locator}: {str(e)}")
            return None

    def switch_to_new_window(self):
        """
        Cambia al control a una nueva ventana/pestaña.
        
        Returns:
            bool: True si el cambio fue exitoso
        """
        try:
            # Obtener el handle original
            original_window = self.driver.current_window_handle
            
            # Esperar nueva ventana
            WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 1)
            
            # Cambiar a la nueva ventana
            for window_handle in self.driver.window_handles:
                if window_handle != original_window:
                    self.driver.switch_to.window(window_handle)
                    logger.info("Cambiado exitosamente a nueva ventana")
                    return True
                    
            return False
            
        except Exception as e:
            logger.error(f"Error cambiando de ventana: {str(e)}")
            return False

    def close_browser(self):
        """Cierra el navegador de forma segura"""
        try:
            if self.driver:
                self.driver.quit()
                logger.info("Navegador cerrado exitosamente")
        except Exception as e:
            logger.error(f"Error cerrando el navegador: {str(e)}")

    def __enter__(self):
        """Soporte para uso con 'with' statement"""
        self.setup_driver()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Cierra el navegador al salir del contexto"""
        self.close_browser()

    def take_screenshot(self, filename):
        """
        Toma una captura de pantalla.
        
        Args:
            filename: Nombre del archivo para guardar la captura
            
        Returns:
            bool: True si la captura fue exitosa
        """
        try:
            self.driver.save_screenshot(filename)
            logger.info(f"Captura de pantalla guardada: {filename}")
            return True
        except Exception as e:
            logger.error(f"Error tomando captura de pantalla: {str(e)}")
            return False

    def clear_cookies(self):
        """Limpia las cookies del navegador"""
        try:
            self.driver.delete_all_cookies()
            logger.info("Cookies limpiadas exitosamente")
        except Exception as e:
            logger.error(f"Error limpiando cookies: {str(e)}")