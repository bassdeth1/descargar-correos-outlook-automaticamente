# Descargar Correos de Outlook Automáticamente

Este repositorio contiene un script en Python creado por **Agustín Arellano, Analista Estadístico**, diseñado para automatizar la descarga de correos desde Outlook. Este sistema permite gestionar de manera eficiente correos electrónicos, filtrando y descargando automáticamente los archivos adjuntos que coincidan con ciertas palabras clave. Este proyecto fue creado con el propósito de optimizar la gestión del flujo de trabajo, ahorrando tiempo y reduciendo errores manuales en el manejo de correos.

## ¿Para qué sirve este sistema?

Este sistema está diseñado para automatizar la descarga de correos electrónicos desde Outlook, permitiendo a los usuarios establecer palabras clave para filtrar correos importantes. Es ideal para aquellos que manejan grandes volúmenes de correos electrónicos y necesitan extraer archivos adjuntos de manera rápida y eficiente, sin necesidad de realizar el proceso manualmente.

### Casos de uso:

- **Manejo de Reportes Automáticos:** Si recibes reportes diarios, semanales o mensuales por correo electrónico, este sistema puede descargarlos automáticamente y almacenarlos en la carpeta que elijas.
- **Filtrado de Correos Importantes:** Solo descargará los correos que contengan ciertas palabras clave, evitando que descargues adjuntos innecesarios.
- **Automatización del Trabajo Diario:** Puedes programar el script para que se ejecute automáticamente cada día, asegurando que siempre tendrás los últimos reportes o documentos sin tener que buscar manualmente en tu bandeja de entrada.

## Ventajas de usar este sistema

- **Ahorro de tiempo:** Automatiza tareas repetitivas de búsqueda y descarga de archivos adjuntos, permitiendo que te concentres en tareas más importantes.
- **Eficiencia:** Evita el error humano en la búsqueda manual de correos electrónicos. El sistema asegura que no se pase por alto ningún correo importante.
- **Facilidad de uso:** Una vez configurado, el sistema funciona de manera autónoma. Solo debes ingresar las palabras clave y la carpeta de destino la primera vez.
- **Integración con Outlook:** Al estar integrado con Outlook, no requiere configuraciones complicadas ni la instalación de nuevos servicios de correo. Utiliza el mismo correo que ya tienes configurado.
- **Personalización:** Puedes ajustar las palabras clave para que el sistema filtre solo los correos relevantes para ti.

## Tecnologías Utilizadas

Este proyecto está basado en Python y utiliza las siguientes bibliotecas:

- **Python 3:** Lenguaje de programación principal utilizado para el script.
- **Win32com:** Para la interacción con Outlook a través de COM, permitiendo acceder y manipular los correos electrónicos directamente desde la aplicación.
- **Psutil:** Para verificar el estado de los procesos de Outlook y asegurarse de que esté en ejecución.
- **Subprocess:** Para abrir Outlook si no está en ejecución y permitir que el script funcione sin intervención manual.

## Requisitos

Antes de ejecutar el script, asegúrate de cumplir con los siguientes requisitos:

### Requisitos del Sistema

- **Sistema Operativo:** Windows (con Outlook instalado y configurado)
- **Python 3.7+** instalado en tu sistema.

### Requisitos de Python

Debes instalar las siguientes bibliotecas de Python ejecutando el siguiente comando:

```bash
pip install -r requirements.txt
