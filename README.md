# Macros para LibreOffice Calc
Este repositorio de macros en Basic (.bas) contiene herramientas para automatizar tareas comunes como la manipulación de fechas, el procesamiento de coordenadas y la adición de marcas de tiempo en tus hojas de cálculo.

| Archivo | Rutina Principal | Descripción Breve |
| :--- | :--- | :--- |
| `fechas.bas` | `AgregarColumnaFechaNumero` | Busca la columna `FECHA_ALTA`, inserta una columna nueva llamada `FECHA` y la llena con la fecha en formato numérico para permitir un ordenamiento correcto. Llama a la subrutina `OrdenarPorFecha` para ordenar la tabla de forma descendente. |
| `utilidades_fecha.bas` |  | Contiene las rutinas auxiliares privadas (`ColumnIndexToLetters` y `OrdenarPorFecha`) que son llamadas desde `AgregarColumnaFechaNumero`. |
| `coordenadas.bas` | `GenerarNubeDePuntos` | Procesa una hoja de datos para filtrar registros por una lista de distritos (`Rafaela`, `Bella Italia` por defecto), formatea una columna de coordenadas (reemplazando comas por tabulaciones y puntos por comas) y vuelca el resultado a un archivo de texto (`coordenadas.txt`) junto con el usuario correspondiente. |
| `reportes.bas` | `AgregarFechaSiVacio` | Recorre las filas de la primera hoja. Si la Columna A (índice 0) tiene contenido y la Columna N (índice 13) está vacía, inserta la fecha actual en la Columna N y le aplica un formato de fecha. |


## Importar macro

Exportar a Hojas de cálculo
Para instalar estos archivos .bas en tu LibreOffice Calc, dirigirse al menú:

`Herramientas → Macros → Organizar macros → Basic...`.

En el panel de la izquierda, seleccionar Mis Macros (para que estén disponibles en cualquier documento) o una librería específica.


## Ejecutar macros

Una vez importadas, las macros estarán disponibles para su ejecución. Dirigirse a:

`Herramientas → Macros → Ejecutar macro...`.

En el listado, navegar a la librería y módulo correspondiente y hacer clic en *Ejecutar*.

> [!NOTE]
> Para la GenerarNubeDePuntos, **desactivar Object Snap** (<kbd>F3</kbd>) previo a seleccionar el archivo de coordenadas para evitar errores de agrupación de círculos.
