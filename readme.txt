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













