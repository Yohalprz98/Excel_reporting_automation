# Automatización de informes de Excel

Este proyecto automatiza el proceso de creación de informes a partir de archivos Excel, generando visualizaciones e integrando datos de múltiples archivos.

## Descripción general

Este script de Python utiliza las bibliotecas `pandas`, `plotly`, y `xlwings` para automatizar la generación de informes a partir de archivos Excel. Realiza las siguientes tareas:

*   Fusiona datos de múltiples archivos Excel en un único libro de Excel de resumen.
*   Copia un rango específico de celdas ("B5").expand() de cada archivo Excel en un libro de resumen.
*   Crea un encabezado estilizado en el libro de resumen.
*   Calcula el total de ventas por país utilizando pandas.
*   Inserta los datos agrupados en el libro de resumen.
*   Crea y añade gráficos de Excel y gráficos de pandas al libro de resumen.
*   Guarda el libro de resumen y cierra la instancia de Excel.

![newplot](https://github.com/user-attachments/assets/a66c4e7b-6f47-4602-b548-4c35b38e9f79)
![image](https://github.com/user-attachments/assets/5bd45d5f-f4b4-4373-a127-5caecbd3ec9e)
