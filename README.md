# Diagramas de Gantt
---
<br/>
<p align="center">
<img src="https://www.sinnaps.com/wp-content/uploads/2017/01/henry-lawrence-gantt.jpg" alt="gantt">
</p>
## Macros para la creaci�n
Con esta macro se pretende facilitar el proceso de creaci�n de Gr�ficos de Gantt. El "Gantt project planner" como plantilla de Excel es aceptable, sin embargo, la macro busca automatizar el proceso de creaci�n y dar mayor flexibilidad en la duraci�n del proyecto.
Integrar la creaci�n autom�tica del gr�fico de Gantt a Excel puede ser muy beneficiosa, ya que en ocaciones es m�s comodo usar todo la informaci�n en una sola plataforma, sumado a que algunos colaboradores no tendran acceso a "Microsoft Project" por ejemplo.
Ya que deben de adquirirlo, aprender a utilizarlo y regularmente se presenta de vuelta en Excel. 
La creaci�n de gr�ficos de Gantt se construye mediante dos macros: Insertar y Grafico_Gantt.
>Macro: Se usa para tareas que se realizan reiteradamente en Excel, con el fin de automatizarlas. Esta puede ser programada mediante el lenuaje de programaci�n Visual Basic.
"Insertar" tiene como fin programar las actividades en d�as, donde se pide el nombre de las tareas, los responsables de cada actividad, el inicio y el fin de las actividades en d�as.
>Gr�fico de Gantt: Es una herramienta que tiene por objetivo expresar el tiempo esperado y real de las actividades de un proyecto. 
"Grafico_Gantt" tiene como objetivo la creaci�n del diagrama. Se puede especificar el t�tulo, lo dem�s estar� dado por la tabla de "Insertar" (macro 1). 
<br/>
<p align="center">
<img src="Diagramas.png" alt="gantt">
</p>

## Criterios a considerar
* El inicio y el fin de las actividades es el tiempo en el cual esperamos (<b>E</b>) realizar la actividad. 
* Las fechas se ingresan en las casillas <b>Inicio</b> y <b>Fin</b> en la forma: dd/mm/aaaa. Sin embargo, funciona bien con n�meros enteros.
* "Insertar" (macro 1) usa formularios para ingresar los datos, al finalizar �nicamente se debe de cerrar esta.
* Una vez creado el diagrama se recomienda actualizar las casillas del tiempo real: <b>R</b>.

## Instalaci�n
El script fue realizado en Microsoft Excel 2016. Se debe insertar como m�dulo dentro del entorno de programaci�n de Excel. Recuerda que son dos scripts los que hay que correr. Primero "Insertar" y despues "Grafico_Gantt".
�Suerte en tu proyecto!
