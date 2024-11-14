18/09/2023

BOT API Telegram para comunicaciones


El bot de Telegram se crea mediante BotFather, propia aplicación de telegram para generar BOTS. 

Nombre:
MRH_Notificaciones

Usuario BOT:
MRH_Notificaciones_bot

Token Access:
5770884977:AAGDI5TJ5vn_e_DYtumrfqrrY6iFX9dClWg

Conecto con el Bot usando la libreria (Paquete Nugget TelegramBots)

13:47. Hago prueba de conexión con el servicio, funciona correctamente

Para poder enviar mensajes por telegram usando el bot, tenemos que utilizar la instrucción:


//Inicialización del componente
Imports Telegram Bot

//En el comienzo de la clase hay que realizar las configuraciones:
Shared Bot_client As TelegramBotClient = New TelegramBotClient("5770884977:AAGDI5TJ5vn_e_DYtumrfqrrY6iFX9dClWg")

Siendo la secuencia alfa-numerica, el token de acceso. 

//Instrucción en envío
 Await Bot_client.SendTextMessageAsync(TelegramID, Mensaje)


Tabla en SAGE para asignar numeros de teléfono con sus IDs de Telegram.
La tabla es:
MRH_TelegramTelefono


Queda funcionando y probado 20/09/2023


//Instrucción de envío de archivos (Hasta 10 MB)

	
	Dim RutaArchivo As String 
	RutaArchivo="Pon aqui la ruta donde se encuentra el archivo que quieres mandar"

	
	'Lo mete en un FileStream ya que el await no se traga como parametro un string del tirón y después lo manda
     Dim InputFileStream As New InputFileStream(New FileStream(RutaArchivo, FileMode.Open), Path.GetFileName(RutaArchivo))
     Await Bot_client.SendDocumentAsync(TelegramID, InputFileStream)

 End If











