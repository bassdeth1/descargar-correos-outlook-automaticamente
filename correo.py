import os
import json
import shutil
import datetime
import psutil  # Para verificar si Outlook está cerrado
import subprocess  # Para abrir Outlook si está cerrado
import win32com.client  # Para interactuar con Outlook a través de COM
from threading import Thread  # Para ejecutar la descarga de correos en un hilo separado
import pythoncom  # Necesario para inicializar COM

# Nombre del archivo JSON para almacenar la configuración
config_file = "config.json"

# Obtener la ruta del script actual
ruta_script = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")

# Función para cargar la configuración desde el archivo JSON
def cargar_configuracion():
    if os.path.exists(config_file):
        with open(config_file, "r") as file:
            return json.load(file)
    return {}

# Función para guardar la configuración en el archivo JSON
def guardar_configuracion(configuracion):
    with open(config_file, "w") as file:
        json.dump(configuracion, file, indent=4)

# Función para solicitar al usuario la ruta y las palabras clave si no están en el JSON
def obtener_ruta_y_palabras_clave():
    configuracion = cargar_configuracion()
    
    if "ruta_trabajo" not in configuracion or not os.path.exists(configuracion["ruta_trabajo"]):
        while True:
            ruta_trabajo = input(f"Por favor, ingresa la ruta de trabajo completa (ej. {ruta_script}): ").strip().replace("\\", "/")
            try:
                if not os.path.exists(ruta_trabajo):
                    print("La ruta no existe. Creándola...")
                    os.makedirs(ruta_trabajo)
                configuracion["ruta_trabajo"] = ruta_trabajo
                break  # Salir del bucle si la ruta es válida
            except PermissionError:
                print(f"Permiso denegado para crear la ruta: {ruta_trabajo}. Por favor, intenta con otra ruta.")
            except OSError as e:
                print(f"Error en la ruta ingresada: {e}. Por favor, intenta con otra ruta.")
    
    if "palabras_clave" not in configuracion or not configuracion["palabras_clave"]:
        palabras_clave = input("Por favor, ingresa una o más palabras clave para buscar en los correos, separadas por comas: ").strip().split(",")
        configuracion["palabras_clave"] = [palabra.strip().lower() for palabra in palabras_clave]  # Guardar en minúsculas y sin espacios adicionales
    
    guardar_configuracion(configuracion)
    return configuracion["ruta_trabajo"], configuracion["palabras_clave"]

# Cargar la ruta de trabajo y las palabras clave desde la configuración
ruta_trabajo, palabras_clave = obtener_ruta_y_palabras_clave()

# Esta función verifica si Outlook está cerrado y lo abre si es necesario. Luego, inicia la descarga de adjuntos en un hilo separado.
def iniciar_descarga_adjuntos():
    try:
        # Verifico si Outlook está cerrado
        outlook_process = "OUTLOOK.EXE"
        if outlook_process not in (p.name() for p in psutil.process_iter()):
            print("Outlook no está abierto. Intentando abrir Outlook...")
            subprocess.Popen(["C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"])  # Abro Outlook

        # Creo y ejecuto el hilo para procesar los correos de Outlook
        print("Iniciando el hilo para procesar correos...")
        thread = Thread(target=procesar_correos_outlook)
        thread.start()
    except Exception as e:
        print(f"Error al iniciar la descarga de adjuntos: {e}")

# Esta función conecta a Outlook y procesa los correos.
def procesar_correos_outlook():
    try:
        # Inicializar el sistema COM
        pythoncom.CoInitialize()
        
        # Conexión a Outlook y acceso a la bandeja de entrada
        print("Conectando a Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
        bandeja_entrada = mapi.GetDefaultFolder(6)  # Carpeta de bandeja de entrada (valor 6)
        
        # Forzar sincronización (si es posible)
        try:
            print("Forzando actualización de correos en Outlook...")
            accounts = mapi.Folders
            for account in accounts:
                account.SendAndReceive(False)
        except Exception as e:
            print(f"No se pudo forzar la actualización: {e}. Continuando con la descarga de correos.")

        # Limpio todos los archivos de la carpeta
        limpiar_carpeta_establecimiento()
        
        print(f"Procesando correos en la carpeta: {bandeja_entrada.Name}...")
        procesar_mensajes_bandeja(bandeja_entrada)
        print("Archivos procesados correctamente")  # Mensaje de confirmación cuando se procesan los archivos
    except Exception as e:
        print(f"Error al procesar correos: {e}")
    finally:
        pythoncom.CoUninitialize()

# Esta función limpia la carpeta de destino eliminando todos los archivos.
def limpiar_carpeta_establecimiento():
    carpeta_establecimiento = os.path.join(ruta_trabajo, "establecimiento").replace("\\", "/")
    print(f"Carpeta de establecimiento: {carpeta_establecimiento}")
    
    if not os.path.exists(carpeta_establecimiento):
        print("La carpeta de establecimiento no existe. Creándola...")
        os.makedirs(carpeta_establecimiento)
    
    for archivo in os.listdir(carpeta_establecimiento):
        ruta_archivo = os.path.join(carpeta_establecimiento, archivo).replace("\\", "/")
        try:
            if os.path.isfile(ruta_archivo):
                os.remove(ruta_archivo)  # Elimino el archivo
                print(f"Archivo eliminado: {ruta_archivo}")
            elif os.path.isdir(ruta_archivo):
                shutil.rmtree(ruta_archivo)  # Elimino la carpeta
                print(f"Carpeta eliminada: {ruta_archivo}")
        except Exception as e:
            print(f"No se pudo eliminar {ruta_archivo}: {e}")  # Capturo cualquier error al eliminar

# Esta función procesa los mensajes en la bandeja de entrada y descarga los adjuntos relevantes.
def procesar_mensajes_bandeja(carpeta):
    try:
        print(f"Revisando mensajes en la carpeta: {carpeta.Name}...")
        # Recorro todos los mensajes en la carpeta de correos
        if hasattr(carpeta, 'Items'):
            mensajes = carpeta.Items
            mensajes.Sort("[ReceivedTime]", True)  # Ordeno los mensajes por fecha de recepción, de más reciente a más antiguo
            
            lista_mensajes = [mensaje for mensaje in mensajes]  # Convierto los mensajes en una lista para evitar errores de índice
            
            for mensaje in lista_mensajes:
                if hasattr(mensaje, 'ReceivedTime'):
                    fecha_mensaje = mensaje.ReceivedTime.date()
                    if fecha_mensaje == datetime.datetime.today().date():
                        print(f"Revisando mensaje recibido el: {fecha_mensaje}")
                        if verificar_palabras_clave(mensaje):
                            # Descargo los adjuntos de los mensajes que coinciden con la fecha de hoy y contienen las palabras clave
                            guardar_adjuntos_mensaje(mensaje)

        # Recorrer las subcarpetas recursivamente para buscar más mensajes
        if hasattr(carpeta, 'Folders'):
            for subcarpeta in carpeta.Folders:
                procesar_mensajes_bandeja(subcarpeta)
    except Exception as e:
        print(f"Error al procesar mensajes en la carpeta {carpeta.Name}: {e}")

# Esta función guarda los adjuntos de un mensaje si cumplen con los criterios.
def guardar_adjuntos_mensaje(mensaje):
    try:
        for adjunto in mensaje.Attachments:
            nombre_adjunto = adjunto.FileName
            if nombre_adjunto.endswith((".xls", ".xlsx", ".xlsm")):
                carpeta_establecimiento = os.path.join(ruta_trabajo, "establecimiento").replace("\\", "/")
                crear_carpeta_si_no_existe(carpeta_establecimiento)
                ruta_guardado = os.path.join(carpeta_establecimiento, nombre_adjunto).replace("\\", "/")
                adjunto.SaveAsFile(ruta_guardado)  # Guardo el adjunto en la carpeta establecida
                print(f"Adjunto guardado: {ruta_guardado}")
    except Exception as e:
        print(f"Error al guardar adjuntos: {e}")

# Esta función crea una carpeta si no existe.
def crear_carpeta_si_no_existe(ruta):
    if not os.path.exists(ruta):
        os.makedirs(ruta)

# Esta función verifica si el asunto o los adjuntos contienen alguna de las palabras clave.
def verificar_palabras_clave(mensaje):
    asunto = mensaje.Subject.lower()
    adjuntos_contienen_palabra_clave = False
    
    for palabra in palabras_clave:
        if palabra in asunto:
            adjuntos_contienen_palabra_clave = True
    
        for adjunto in mensaje.Attachments:
            nombre_adjunto = adjunto.FileName.lower()
            if nombre_adjunto.endswith((".xls", ".xlsx", ".xlsm")) and (palabra in nombre_adjunto):
                adjuntos_contienen_palabra_clave = True
    
    return adjuntos_contienen_palabra_clave

# Inicia el proceso cuando se ejecuta el script
iniciar_descarga_adjuntos()
