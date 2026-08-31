# Asistente de Automatización de RMA (RPA Bot) 🤖🚀

## 📝 Descripción del Proyecto
Este proyecto es una herramienta de Automatización de Procesos Robóticos (RPA) desarrollada en **Python** para entornos Windows. Su objetivo principal es optimizar y agilizar la introducción de datos en el sistema de gestión corporativo de RMA, reduciendo drásticamente los tiempos de procesamiento de información y eliminando errores operativos.

El asistente implementa una interfaz gráfica (GUI) amigable que se conecta con la suite de Microsoft Office y el navegador Microsoft Edge para capturar, procesar e inyectar datos de reparaciones de forma automatizada.

---

## 🛠️ Tecnologías y Librerías Utilizadas
- **Lenguaje Principal:** Python 🐍
- **Interfaz Gráfica (GUI):** Tkinter (Gestión dinámica de ventanas y formularios operativos).
- **Integración Office (COM):** `pywin32` (`win32com.client`) para la comunicación e interoperabilidad con la API nativa de Microsoft Outlook.
- **Procesamiento de Textos:** `re` (Expresiones Regulares avanzadas para la minería de datos y filtrado de strings).
- **Automatización RPA:** `pyautogui` / `pynput` (Simulación y escucha de periféricos como teclado, ratón y captura precisa de posiciones físicas en pantalla).
- **Persistencia de Datos:** `json` (Estructuración del mapeo de pantallas mediante archivos de configuración).
- **Herramientas de Portapapeles:** `pyperclip` (Inyección fluida de texto mediante memoria temporal).

---

## ⚙️ Funcionalidades Clave

1. **Minería de Datos e Integración con Outlook (Paso 1):** 
   - El script interactúa en tiempo real con la ventana activa de Outlook (`ActiveExplorer`).
   - Utiliza expresiones regulares dinámicas para escanear el cuerpo y el asunto del correo electrónico seleccionado.
   - Extrae de forma inteligente campos críticos: Código de modelo (con patrones específicos como `HA-`, `CS-`, `KW-`), Números de Serie (`S/N`) y Número de RMA.
   - Cuenta con un algoritmo de filtrado secuencial (de arriba a abajo) para capturar la descripción de la avería, aislando palabras clave operativas (`no funciona`, `error`, `falla`) y activando un cortafuegos de texto para ignorar firmas y avisos legales.

2. **Inyección de Datos y Control de Periféricos (Paso 2):** 
   - Automatiza la suite de clics e inserción de caracteres en el formulario web de Microsoft Edge.
   - Centraliza los datos estáticos repetitivos en el código (técnico asignado, marcas habituales) permitiendo que un operador realice modificaciones de mantenimiento sencillas si los parámetros de la empresa cambian.

3. **Sistema de Calibración por Coordenadas (JSON):** 
   - Integra un asistente guiado que utiliza un `Listener` en segundo plano para escuchar eventos del teclado. 
   - Permite al usuario mapear los píxeles (X, Y) de cualquier monitor simplemente posicionando el ratón y pulsando la tecla `CTRL`. 
   - Exporta estas posiciones de forma dinámica a `coordenadas_bot.json`, dotando al software de portabilidad total ante cualquier resolución o cambio en el diseño de la web.

4. **Automatización del Flujo de Trabajo Diario:** 
   - Gestión inteligente del flujo del calendario y fecha en curso que optimiza la inicialización diaria del sistema de gestión mediante un único clic obligatorio guiado por el ratón.

---

## 📂 Estructura del Repositorio
- `robot_rma.py`: Script de lógica principal, algoritmos de extracción RegEx, interfaz Tkinter y motor RPA.
- `coordenadas_bot.json`: Archivo de configuración local donde se guarda la persistencia de los píxeles de pantalla calibrados.
- `ejecutar_robot_rma.bat`: Script de inicialización por lotes (`@echo off`) diseñado para verificar e instalar las dependencias automáticamente vía `pip` y lanzar el proceso ocultando la consola (`pythonw`) para mejorar la experiencia de usuario.

---
🔧 *Proyecto desarrollado de forma autónoma enfocado en la ingeniería de software pragmática, la automatización del puesto de trabajo (RPA) y la optimización de procesos informáticos.*
