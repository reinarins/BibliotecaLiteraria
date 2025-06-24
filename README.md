# 📚 Mini Biblioteca de Alejandría

## Descripción del proyecto
Este es un proyecto que simula una mini biblioteca y en el que he integrado tres tecnologías:
- <b>Base de datos en SQL Server</b> (SQL Server Management Studio), cargado  mediante un notebook 'biblioteca-setup.ipynb'
- <b>Robot RPA diseñado en UiPath Studio</b>, conectado a la base de datos para generar informes Excel anuales
- <b>Dashboard en Power BI</b>, realizado para visualizar estadísticas clave

## Objetivos del proyecto
1. <b>Crear una base de datos</b> sencilla y sólida que proporcione los siguientes datos de cada libro:
  - Título
  - Autor/a
  - Editorial
  - Género
  - Año de publicación

  Y que, además, registre las fechas en las que se ha proporcionado cada libro a usuarios/as de la biblioteca, en calidad de préstamo. Las fechas se comprenden entre los años 2016 y 2024, y la cantidad de préstamos por libro se ha establecido acorde a los ránkings proporcionados por el blog 'Comunidad Baratz'. Para 2024, se ha realizado una estimación debido a que los libros más prestados están divididos por trimestres.

2. <b>Diseñar un flujo automatizado con UiPath</b> que se conecte a la base de datos, y que, a partir de los datos obtenidos, genere reportes anuales en un archivo Excel llamado 'histórico_préstamos', con las siguientes columnas:
  - Título
  - Autor/a
  - Editorial
  - Género
  - Publicación
  - Préstamos

  Los títulos están ordenados según su número de préstamos, en orden descendente.

3. <b>Mejorar habilidades en Power BI</b> para la visualización de datos.

## Tecnologías utilizadas
- SQL Server Management Studio (SSMS)
- UiPath Studio
- Power BI

## Origen de los datos
- Datos de los libros más prestados en las bibliotecas públicas españolas entre 2016 y 2023, extraídos del blog 'Comunidad Baratz' (https://www.comunidadbaratz.com)
- Datos de los libros más prestados en la red de bibiotecas de Madrid durante 2024 (https://www.comunidad.madrid/portal-lector/bibliotecas/bibliotecas-red/titulos-prestados)

## Capturas de pantalla
![image](https://github.com/user-attachments/assets/06ca5b21-111d-4060-88f2-3f0110a5f9fc)
<p align="right"><small><i>Consulta SQL en la base de datos: Títulos publicados en 2019</i></small></p>

![image](https://github.com/user-attachments/assets/f5f009b5-acd7-4585-a764-eba8e0815459)
<p align="right"><small><i>Consulta SQL en la base de datos: Títulos prestados en 2024</i></small></p>

![image](https://github.com/user-attachments/assets/25ea6c65-2b3d-4905-ab26-f1f2e6252d4f)
<p align="right"><small><i>Conexión del robot (creado en UiPath Studio) con la base de datos, dentro de una actividad 'Try Catch' para el manejo de errores (por ejemplo, en caso de que no pueda establecerse la conexión con la base de datos)</i></small></p>

![image](https://github.com/user-attachments/assets/8ed2f136-85a1-40f1-8b14-e5d6700d2d7b)
<p align="right"><small><i>Se realiza la consulta mediante una actividad 'Run Query'. Así, obtenemos el archivo Excel con los datos de cada año recogidos en su hoja correspondiente</i></small></p>

![image](https://github.com/user-attachments/assets/75209d28-fad3-446b-9daa-cac7ded5da68)
<p align="right"><small><i>Consulta SQL desde UiPath Studio</i></small></p>


![image](https://github.com/user-attachments/assets/8eb41bed-eb7d-4dab-b53f-a7636ef0a958)
<p align="right"><small><i>Se vuelcan los datos en el Excel y se ejecuta la macro importada previamente. En este caso, he optado por crear la macro previamente e importarla en el Excel para evitar problemas de seguridad a la hora de otorgar permisos que pudieran ejecutar código peligroso por error</i></small></p>

![image](https://github.com/user-attachments/assets/e32ae4e1-b25b-4395-8585-e2def2f1d398)
<p align="right"><small><i>Macro en VBA utilizada para aplicar un formato automático a los informes generados por el robot. Ajusta anchos de columna, aplica un estilo "cebra" para mejorar la legibilidad entre filas, y destaca bordes y encabezados</i></small></p>

![image](https://github.com/user-attachments/assets/3fa25461-96e8-41a9-9120-efed615ca522)
<p align="right"><small><i>Informe creado en Power BI</i></small></p>

## Cómo usar

1. Restaurar la base de datos desde 'database/biblioteca-setup.ipynb'
2. Abrir el proyecto de UiPath desde 'uipath-robot/'
3. Ejecutar el robot para generar los reportes
4. Visualizar los resultados en Power BI ('powerbi/dashboard.pbit')

## Futuras mejoras
- Automatizar la actualización periódica en Power BI
- Incluir más fuentes de datos 💻📚

## Autora
Desarrollado por <a href="https://github.com/reinarins">@reinarins</a>
<br>
Desarrolladora RPA | UiPath | Innovación digital desde una perspectiva humana y ética
<br>
Conecta conmigo en <a href="https://www.linkedin.com/in/irisfrro/">LinkedIn</a>!
