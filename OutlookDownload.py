import win32com.client
import os
from datetime import datetime, timedelta

#Creamos nuestro objeto que representa la aplicacion que deseamos manipular
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#Accedemos a la bandeja de entrada (GetDefaultFolder(6))
bandeja_entrada = outlook.GetDefaultFolder(6)

#Obtenemos los mensajes que se recibieron desde la fecha indicada hasta la fecha acutal
fecha_ayer = datetime.now() - timedelta(days=1)
fecha_ayer = fecha_ayer.strftime('%d/%m/%y')

mensajes = bandeja_entrada.Items.restrict(f'[SentOn] > "{fecha_ayer}"')

#Ordenamos los mensajes segun el tiempo en el que se recibieron
mensajes.Sort("[ReceivedTime]", True)

print('\n\n'+'Procesando lectura de correos para ubicar archivos adjuntos...'.upper().center(85)+'\n')

#Iteramos sobre cada mensaje una vez que han sido ordenados segun el tiempo en que se recibieron
for mensaje in mensajes:
    subject_mensaje=  mensaje.subject#obtenemos el subject del mensaje

    #Si el string utilizado en el condicional se encuentre el string que hemos especificado 
    #entonces descargamos el archivo adjunto que este contiene.
    s1 = 'Resultado Reporte de Actividades VM'#Subject que utiliza Hanliet
    s2 = 'Resultado de Actividades VM'#Subject que utiliza Jose Pablo
    s3 = 'Resultados de Actividades VM'#Subject que utiliza Lorenz

    if s1 in subject_mensaje or s2 in subject_mensaje or s3 in subject_mensaje:

        try:
            #El metodo Attachments nos retorna una lista de objetos que representan los archivos
            #adjuntos,es necesario obtener el listado de archivos adjuntos ya que en los correos
            #concatenados podrian venir imagenes adjuntas por lo tanto debemos revisarlos todos para
            #identificar cual o cuales archivos nos interesan descargar  
            adjuntos = mensaje.Attachments
            #iteramos los ojetos que representan los archivos adjuntos y usamos su metodo FileName para obtener el nombre
            for adjunto in adjuntos:
                if 'Actividades VM' in adjunto.FileName:
                    nombre_archivo_adjunto = adjunto.FileName
                    break

            #Guardamos en una variable el path donde guardaremos el archivo adjunto
            directorio_destino = f'D:/Soporte_Core_CA/Gestion de Cambios/Resultado_Actividades/{nombre_archivo_adjunto}'
            #Descargamos el archivo adjunto y lo guardamos en el path especificado en la variable directorio_destino
            adjunto.SaveASFile(directorio_destino)
            #Imprimimos informacion relevante por pantalla 
            print(f'\nSUBJECT DEL CORREO: {mensaje.subject}')
            print(f'\nNOMBRE DEL ARCHIVO ADJUNTO: {nombre_archivo_adjunto}')
            print(f'\n\n'+'El archivo ha sido guardado con exito!!!'.upper().center(80))
            break #Finalizamos el for para que no siga iterando sobre mas correos puesto que ya obtuvimos lo que necesitabamos

        except Exception as e:
            print('No se pudo guardar el archivo adjunto por la siguiente razon:\n')
            print(e)    

#Dejamos este input para poder ver lo que ocurrio y tener el control de 
# cuando cerrar el programa.
input()
