# XM_Pub_Checker
 
XM opera el Sistema Interconectado Nacional (SIN) y administramos el Mercado de Energía Mayorista (MEM), para lo cual realiza las funciones de Centro Nacional de Despacho -CND-, Administrador del Sistema de Intercambios Comerciales -ASIC- y Liquidador de Cuentas de Cuentas de cargos por Uso de las redes del Sistema Interconectado Nacional - LAC. Además, XM administra las Transacciones Internacionales de Electricidad de corto plazo -TIE- con Ecuador ([enlace](https://www.xm.com.co/corporativo/Paginas/Nuestra-empresa/que-hacemos.aspx) para mayor informacion sobre XM).

Entre una de las tantas actividades que realiza XM, se encuentran las publicaciones de Registros y Fallas diarios de fronteras comerciales, estas se pueden encontrar ingresando en la pagina de XM en el siguiente [enlace](https://www.xm.com.co/Paginas/Mercado-de-energia/fronteras-comerciales-en-proceso-de-registro.aspx). Este es un proceso que para los representantes de fronteras y operadores de red es de gran intereses, pero es un trabajo el cual se vuelve tedioso ya que las publicaciones no son siempre en la misma hora, y si se quiere tener un control diario sobre estas publicaciones el dejar encargada a una persona su revision, puede ser muy poco productivo.

Con el fin de generar una tarea automatica, eficiente y eficaz con respecto a las publicaciones que realiza XM, se diseño el siguiente codigo, el cual puede ser utilizado por cualquier empresa interesada en automatizar esta tarea.

## Librerias utilizadas

- **datetime:** Para tomar la fecha actual automaticamente.
- **pandas:** Para el manejo de archivos de multiples fuentes.
- **yagmail:** Facilita el uso de gmail para el envio de correos electronicos.
- **os.path:** Para verificar archivos en una ruta.

Se recomienda el uso de anaconda para la creacion de un ambiente virtual.

## Modo de uso

(Esta es una primera version que puede ser actualizada a un encapsulamiento el cual no requiera abrir el codigo para modificar los valores de algunas variables)

Una vez instaladas las librerias necesarias, debemos ingresar al codigo de python para modificar los valores de las variables ```mail_pass, mail_user, mail_2_send```, estas llevan la informacion del correo que se va a utilizar para el envio de mensajes y tambien el correo al cual se debe enviar la informacion.

Con los electronicos definidos, podemos ejecutar el codigo y verificar que todo funcione correctamente, los archivos enviados son almacenados en las carpetas Fallas, Registros, Resumenes, esto con el fin de llevar control de las publicaciones que se han enviado por correo electronico. Pero esto aun no es una tarea automatica, para no complicar las cosas
vamos a utlizador el programador de tareas de windows junto a un archivo bat para lograr este cometido.

### Archivo bat

Para genenar el archivo bat vamos a utilizar un archivo de texto y vamos a seguir el ejemplo dentro del archivo **executeF3.txt** con los siguiente comandos:
 
```Python
@echo on
cd "C:\Python Scripts\Bot_Publicaciones_XM"
call C:\ProgramData\Anaconda3\Scripts\activate.bat
call conda activate botXM
call "C:\Users\gcoral\.conda\envs\botXM\python.exe" bot_Xm_pub.py
```

Este es un ejemplo de como deberia ser el archivo bat, la primera linea nos ubicamos en la ruta donde se encuentra el codigo python, posteriormente activamos el CMD de conda para asi activar nuestro ambiente con las librerias instaladas y finalmente ejecutar el codigo. Una vez tengamos listo nuestro archivo, simplemente modificamos la extension de este de .txt a .bat.

### Programador de tareas

Este programo es muy sencillo de usar para la programacion de tareas en windows, solo debemos seguir los siguientes pasos:

- Se abre el programador de tarea.
- En la ventana de acciones seleccionamos "Crear tarea"/"Create Task".
- Se abrira una nueva ventana, abrimos la pestaña "Acciones"/"Actions".
- Damos al boton "Nuevo"/"New", seleccionamos la accion "Ejecutar programa"/"Strat program" y en la barra del programa pondremos la ruta de nuestro archivo .bat.
- Ingresamos a la pesta;a "Accionadores"/"Triggers".
- Damos al boton "Nuevo"/"New", seleccionamos la opcion "Diariamente"/"Daily" y agregamos a gusto el numero de veces que queremos que el codigo se ejecute en el dia.

![alt-text](https://github.com/gabrielcoralc/XM_Pub_Checker/blob/main/Task_scheduler.PNG)

## Resultado

Si hacemos todo de manera correcta, deberian empezar a llegar nuestros correo de manera automatica los siguientes mensajes:

![alt-text](https://github.com/gabrielcoralc/XM_Pub_Checker/blob/main/mail_example.PNG)
