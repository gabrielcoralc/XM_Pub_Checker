# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 14:59:58 2021

@author: Gabri
"""

import datetime
from datetime import datetime as dtime
import pandas as pd
import os.path
import yagmail

##Funcion que revisa i
def check_today_data(mail_user,mail_pass,mail_2_send,tipo_pub="Registros",operadores_de_red=["CENTRALES ELECTRICAS DE NARIÑO S.A. E.S.P. - DISTRIBUIDOR"],today_date=None):
    """La funcion ingresa al portal de XM y revisa si la publicacion del dia que se esta buscando ya se encuentre en la pagina.
    Si la publicacion esta en la pagina, se descarga el archivo Excel, filtra los operadores de red de interes, envia la informacion
    al correo que se designe en la funcion send_mail. Finalmente el programa guarda en una carpeta el archivo que se envio,
    para tener constancia de que el archivo ya fue enviado y no volver a realizar el mismo proceso con esta publicacion
    
    Parameters
    ----------
    mail_user : str
        Correo electronico de gmail a utilizar
    mail_pass : str
        Clave del correo electronico a utilizar
    mail_2_send: str
        Correo electronico al cual se envia la informacion
    tipo_pub : str
        Define el tipo de publicacion a descargar, puede ser Fallas o Registros
    operadores_de_red : list
        Lista de cadenas de texto, con los nombre de los operadores de red que se deben filtrar
    today_date: datetime
        Se debe ingresar el un string en formato d/m/Y, el cual buscara el archivo publicado en esa fechar
    Returns
    -------
    La funcion no retorna ningun valor, ejecuta el proceso de enviar por correo y guardar la informacion si es exitoso
    """
    if tipo_pub== "Registros" or tipo_pub== "Fallas":
        print("Buen dia \nSe esta iniciado el proceso de revision de publicaciones en XM\n")
        #Revisamos la fecha y la adecuamos la informacion al formato que necesitamos
        if today_date is None:
            today_date=datetime.date.today()
        else:
            today_date=dtime.strptime(today_date,"%d-%m-%Y")
        if today_date.month<10:
            mes="0"+str(today_date.month)
        else:
            mes=+str(today_date.month)
        if today_date.day<10:
            dia="0"+str(today_date.day)
        else:
            dia=str(today_date.day)
            
        #Vemos si el archivo ya fue creado y por ende enviado
        filename= tipo_pub + "_" + today_date.strftime("%d_%m_%Y") + ".xlsx"  
        if os.path.isfile(tipo_pub+"/"+filename) is False:
            try:
                if tipo_pub== "Registros":
                    df=pd.read_excel(f"https://www.xm.com.co/pubregistros/PubFC{today_date.year}-"+mes+"-"+dia+".xlsx")
                else:
                    df=pd.read_excel(f"https://www.xm.com.co/pubfallahurto/PubFC_Falla-Hurto{today_date.year}-"+mes+"-"+dia+".xlsx")
    
                df=df[df["Operador de Red"].isin(operadores_de_red)]
                df.to_excel(tipo_pub+"/"+filename,index=False)
                send_mail(mail_user,mail_pass,mail_2_send,tipo_pub,operadores_de_red,df,filename,today_date)
                #Se descarga la informacion desde XM, se filtra la informacion segun operador de red y se envia correo al CGM
            except:
                print("No hay conexion a la pagina web de XM O el archivo aun no se ha publicado")
        else:
            print("Archivo ya existe")
        #Vemos si es el inicio de otro mes para enviar resumen del mes anterior
        if today_date.day == 1:
            filename="Resumen_" + tipo_pub + f"_mes_{today_date.month-1}" + ".xlsx"  
            if os.path.isfile("Resumenes/"+filename) is False:
                    df=resumen_reg(month=today_date.month-1, year=today_date.year,operadores_de_red=operadores_de_red,tipo_pub=tipo_pub)
                    df.to_excel("Resumenes/"+filename)
                    send_mail(mail_user,mail_pass,mail_2_send,tipo_pub,operadores_de_red,df,filename,today_date,resumen=True)
    else:
        print("No se ingreso un tipo de publicacion valido.\nSe debe ingresar el valor de 'Fallas' o 'Registros' a la funcion.")


        
def send_mail(mail_user,mail_pass,mail_2_send,tipo_pub,operadores_de_red,df,filename,today_date,resumen=False):
    """La funcion crea el correo que se va a enviar con un resumen de la informacion encontrada
    
    Parameters
    ----------
    mail_user : str
        Correo electronico de gmail a utilizar
    mail_pass : str
        Clave del correo electronico a utilizar
    mail_2_send: str
        Correo electronico al cual se envia la informacion
    tipo_pub : str
        Define el tipo de publicacion a descargar, puede ser Fallas o Registros
    operadores_de_red : list
        Lista de cadenas de texto, con los nombre de los operadores de red que se deben filtrar
    df: dataframe
        Contiene la publicacion descargada
    filename: str
        Nombre con el cual se va a guardar el archivo
    today_date: datetime
        Se debe ingresar el un string en formato d/m/Y, el cual buscara el archivo publicado en esa fechar
    resumen: bool
        Valor que confirma si se debe hacer un resumen del mes de todos los archivos publicados
    Returns
    -------
    La funcion no retorna ningun valor, ejecuta el proceso de enviar por correo
    """
    yag = yagmail.SMTP(user=mail_user, password=mail_pass)
    operadores_string=""
    for operador in operadores_de_red:
        operadores_string+=f"""
        - {operador}"""  
    if resumen==True:
            agentes=df[["Agente que Solicita el Registro" if tipo_pub=="Registros" else "Agente Representante"][0]].unique().tolist()
            agentes_string=""
            for agente in agentes:
                agentes_string+=f"""
                - {agente}"""
            observaciones=df[["Tipo de Solicitud" if tipo_pub=="Registros" else "Equipo en Falla"][0]].unique().tolist()
            observacion_string=""
            for observacion in observaciones:
                observacion_string+=f"""
                - {observacion}"""
            Message=f"""Cordial saludo ingenieros,   

                        Se han detectado {df.shape[0]} {tipo_pub} de fronteras para el Resumen del mes anterior, segun los siguientes Operadores de red filtrados:
                        {operadores_string}

                        Los Agentes representantes de las fronteras que reportan {tipo_pub} son:
                        {agentes_string}
                        
                        Se presentan las siguientes observaciones por {tipo_pub}:
                        {observacion_string}

                        Se adjunta archivo excel con la informacion mencionada para su revision.

                        Gracias por su atencion y les deseo un buen dia,
                     """
            yag.send(to=send_mail, subject="Resumen de "+ f'{tipo_pub} de fronteras para el mes '+str(today_date.month-1), contents=Message,attachments="Resumenes/"+filename)
            print("Summary sent succesfully")
    else:
        try:
            if df.shape[0]>0:          
                    agentes=df[["Agente que Solicita el Registro" if tipo_pub=="Registros" else "Agente Representante"][0]].unique().tolist()
                    agentes_string=""
                    for agente in agentes:
                        agentes_string+=f"""
                        - {agente}"""
                    observaciones=df[["Tipo de Solicitud" if tipo_pub=="Registros" else "Equipo en Falla"][0]].unique().tolist()
                    observacion_string=""
                    for observacion in observaciones:
                        observacion_string+=f"""
                        - {observacion}"""
                    Message=f"""Cordial saludo ingenieros,   

                                Se han detectado {df.shape[0]} {tipo_pub} de fronteras para el dia hoy, segun los siguientes Operadores de red filtrados:
                                {operadores_string}

                                Los Agentes representantes de las fronteras que reportan {tipo_pub} son:
                                {agentes_string}
                                
                                Se presentan las siguientes observaciones por {tipo_pub}:                                
                                {observacion_string}

                                Se adjunta archivo excel con la informacion mencionada para su revision.

                                Gracias por su atencion y les deseo un buen dia,
                             """

                    #sending the email
                    yag.send(to=send_mail, subject=f'{tipo_pub} de fronteras '+today_date.strftime("%d/%m/%Y"), contents=Message,attachments=tipo_pub+"/"+filename)
            else:
                    Message=f"""Cordial saludo ingenieros,

                                Para el dia de hoy no se han encontrado {tipo_pub} publicados por parte de XM, segun los siguientes Operadores de red filtrados:
                                {operadores_string}

                                Gracias por su atencion y les deseo un buen dia,"""
                    yag.send(to=send_mail, subject=f'{tipo_pub} de fronteras '+today_date.strftime("%d/%m/%Y"), contents=Message)
            print("Today report sent successfully")
        except:
            print("Error: Email cant be sent")
            os.remove(tipo_pub+"/"+filename)

def resumen_reg(month,year,operadores_de_red=["CENTRALES ELECTRICAS DE NARIÑO S.A. E.S.P. - DISTRIBUIDOR"],tipo_pub="Registros"):
    """La funcion hace una recopilacion de todos los archivos que fueron publicados en el mes y lo guarda en un df
    
    Parameters
    ----------
    month : int
        numero del mes que se va a hacer el resumen
    year : int
        Numero del año
    operadores_de_red : list
        Lista de cadenas de texto, con los nombre de los operadores de red que se deben filtrar
    tipo_pub : str
        Define el tipo de publicacion a descargar, puede ser Fallas o Registros
    Returns
    -------
    prov_df: dataframe
    
    Dataframe con toda la informacion recopilada de las publicaciones
    """
    day=2
    if month<10:
        month="0"+str(month)
    else:
        month=str(month)
    if tipo_pub== "Registros":
        prov_df=pd.read_excel(f"https://www.xm.com.co/pubregistros/PubFC{year}-"+month+"-"+"01"+".xlsx")
    else:
        prov_df=pd.read_excel(f"https://www.xm.com.co/pubfallahurto/PubFC_Falla-Hurto{year}-"+month+"-"+"01"+".xlsx")
    prov_df=prov_df[prov_df["Operador de Red"].isin(operadores_de_red)]
    while True:
        if day<10:
            day_s="0"+str(day)
        else:
            day_s=str(day)
        try:
            if tipo_pub== "Registros":
                df=pd.read_excel(f"https://www.xm.com.co/pubregistros/PubFC{year}-"+month+"-"+day_s+".xlsx")
            else:
                df=pd.read_excel(f"https://www.xm.com.co/pubfallahurto/PubFC_Falla-Hurto{year}-"+month+"-"+day_s+".xlsx")
            df=df[df["Operador de Red"].isin(operadores_de_red)]
            prov_df=prov_df.append(df,ignore_index=True)
            day=day+1
        except:
            break
    if tipo_pub== "Registros":
        prov_df=prov_df.drop_duplicates(subset=["Código SIC","Tipo de Solicitud","Estado"])
    else:
        prov_df=prov_df.drop_duplicates(subset=["Código SIC"])
    return prov_df
   
if __name__=="__main__":
    # Poner los datos con el correo que se va a utilizar para enviar correo
    # y el correo al que se va a enviar.
    ## Advertencia: el correo electronico que se va a utilizar para enviar mensajes debe estar habilitado
    ## Segun como se indica en las intrucciones de la documentacion de YAGMAIL, esto puede generar problemas de seguridad
    ## en el correo, por lo que es recomendable crear uno nuevo.
    mail_2_send='example@gmail.com.co'
    mail_user='example2@gmail.com'
    mail_pass="PASS_EXAMPLE"
    check_today_data(mail_user,mail_pass,mail_2_send,tipo_pub="Fallas")
    check_today_data(mail_user,mail_pass,mail_2_send,tipo_pub="Registros")