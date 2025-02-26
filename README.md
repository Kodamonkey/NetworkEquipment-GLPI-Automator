# NetworkEquipment-GLPI-Automator

NetworkEquipment-GLPI-Automator es una aplicación para gestionar equipos de red utilizando GLPI y Excel. Permite registrar, actualizar y sincronizar equipos de red entre GLPI y un archivo Excel.

## Requisitos

- Python 3.7 o superior
- pip (gestor de paquetes de Python)

## Instalación

1. Clona el repositorio o descarga los archivos del proyecto.

2. Navega al directorio del proyecto:

    ```sh
    cd /path/to/NetworkEquipment-GLPI-Automator
    ```

3. Instala las dependencias:

    ```sh
    pip install -r requirements.txt
    ```

## Configuración

1. Crea un archivo [.env](http://_vscodecontentref_/0) en el directorio del proyecto con las siguientes variables:

    ```env
    GLPI_URL=http://your-glpi-url
    USER_TOKEN=your-user-token
    APP_TOKEN=your-app-token
    PATH_EXCEL_NETWORK=path/to/network_equipment.xlsx
    IP_CAM_URL=http://your-ip-cam-url
    ```

2. Obtén las variables [GLPI_URL](http://_vscodecontentref_/1), _TOKEN`USER` y [APP_TOKEN](http://_vscodecontentref_/2) desde GLPI:

    - **GLPI_URL**: Es la URL base de tu instancia de GLPI. Por ejemplo, `http://localhost/glpi` o `http://your-glpi-domain`:
        1. Inicia sesión en tu instancia de GLPI como administrador.
        2. Ve a `Setup` > `General` > `API`.
        3. Copia `URL of the API`.

    - **USER_TOKEN**: Para obtener el token de usuario, sigue estos pasos:
        1. Inicia sesión en tu instancia de GLPI.
        2. Ve a `My Settings` (usualmente accesible desde la esquina superior derecha).
        3. En la sección `Remote access keys`, genera un nuevo `API token` si no tienes uno. Copia el token generado.

    - **APP_TOKEN**: Para obtener el token de la aplicación, sigue estos pasos:
        1. Inicia sesión en tu instancia de GLPI como administrador.
        2. Ve a `Setup` > `General` > `API`.
        3. En la sección final, presiona `Add API client` y genera un nuevo token para tu aplicación. Copia el token generado.

3. Asegúrate de que el archivo Excel especificado en [PATH_EXCEL_NETWORK](http://_vscodecontentref_/3) exista. Si no existe, la aplicación lo creará automáticamente.

## Uso

1. Ejecuta la aplicación:

    ```sh
    python app.py
    ```

2. La interfaz gráfica de usuario (GUI) se abrirá. Desde allí, puedes realizar las siguientes acciones:

    - Registrar equipos de red en Excel y GLPI.
    - Sincronizar datos entre Excel y GLPI.
    - Escanear códigos QR para registrar equipos de red.

## Dependencias

Las principales dependencias del proyecto son:

- [tkinter](http://_vscodecontentref_/4): Biblioteca estándar de Python para interfaces gráficas.
- [pandas](http://_vscodecontentref_/5): Biblioteca para manipulación y análisis de datos.
- `opencv-python`: Biblioteca para procesamiento de imágenes y captura de video.
- [pyzbar](http://_vscodecontentref_/6): Biblioteca para decodificación de códigos QR.
- [requests](http://_vscodecontentref_/7): Biblioteca para realizar solicitudes HTTP.
- `python-dotenv`: Biblioteca para cargar variables de entorno desde un archivo [.env](http://_vscodecontentref_/8).
- [urllib3](http://_vscodecontentref_/9): Biblioteca para manejar solicitudes HTTP.
- [numpy](http://_vscodecontentref_/10): Biblioteca para computación numérica.
- [openpyxl](http://_vscodecontentref_/11): Biblioteca para leer y escribir archivos Excel.