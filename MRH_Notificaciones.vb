Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Telegram.Bot

Public Class MRH_Notificaciones
    Shared Bot_client As TelegramBotClient = New TelegramBotClient("5770884977:AAGDI5TJ5vn_e_DYtumrfqrrY6iFX9dClWg")
    Dim TimerLoop As New System.Timers.Timer 'TimerLoop de Correo
    Dim TimerLoopTelegram As New System.Timers.Timer 'TimerLoop de Telegram
    Dim ControladorServicio As New System.ServiceProcess.ServiceController

    'Timer Del envío de Notificaciones por Correo
    Protected Overrides Sub OnStart(ByVal args() As String)

        AddHandler TimerLoop.Elapsed, AddressOf TimerLoop_Tick
        TimerLoop.AutoReset = True
        TimerLoop.Interval = 10000
        TimerLoop.Start()

        AddHandler TimerLoop.Elapsed, AddressOf TimerLoop_Tick2
        TimerLoopTelegram.AutoReset = True
        TimerLoopTelegram.Interval = 7500
        TimerLoopTelegram.Start()

    End Sub


    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

    Private Async Sub TimerLoop_Tick(sender As Object, e As EventArgs)
        Dim conexion As SqlConnection
        Dim sqlcomando As SqlCommand
        Dim consulta As String
        Dim respuesta As IDataReader
        Dim MovPosicion As Guid
        Dim Email As String = ""
        Dim MRH_OrigenNotificacion As String = ""
        Dim Asunto As String = ""
        Dim Mensaje As String = ""
        Dim Adjunto As String = ""
        Dim Adjunto1 As String = ""
        Dim Adjunto2 As String = ""
        Dim Adjunto3 As String = ""
        Dim Adjunto4 As String = ""
        Dim Adjunto5 As String = ""
        Dim Adjunto6 As String = ""
        Dim Adjunto7 As String = ""
        Dim Adjunto8 As String = ""
        Dim Adjunto9 As String = ""
        Dim Adjunto10 As String = ""
        Dim Adjunto11 As String = ""
        Dim Adjunto12 As String = ""
        Dim Adjunto13 As String = ""
        Dim Adjunto14 As String = ""
        Dim Adjunto15 As String = ""
        Dim Adjunto16 As String = ""
        Dim Adjunto17 As String = ""
        Dim Adjunto18 As String = ""
        Dim Adjunto19 As String = ""
        Dim Adjunto20 As String = ""
        Dim Adjunto21 As String = ""
        Dim Adjunto22 As String = ""
        Dim Adjunto23 As String = ""
        Dim Adjunto24 As String = ""
        Dim Adjunto25 As String = ""
        Dim MRH_interno As Integer = 0
        Dim EmailEmisor As String = ""
        Dim PassEmisor As String = ""
        Dim SmtpEmisor As String = ""
        Dim FirmaEmisor As String = ""
        Dim cabeceraMensaje As String = ""
        Dim cuerpoMensaje As String = ""
        Dim pieMensaje As String = ""
        Dim htmlBody As String = ""
        Dim avHtml As AlternateView
        Dim pic1 As LinkedResource
        Dim textBody As String = ""
        Dim TelegramID As String = ""
        Dim m As MailMessage
        Dim client As SmtpClient

        'Funciones.EnviaNotificaciones()
        Try
            conexion = New SqlConnection(Funciones.CadenaConexionSage)
            conexion.Open()
            consulta = "SELECT TOP 1 MovPosicion,Email,Asunto,Mensaje,
                        ISNULL(Adjunto,'') As Adjunto,
                        ISNULL(Adjunto1,'') As Adjunto1,
                        ISNULL(Adjunto2,'') As Adjunto2,
                        ISNULL(Adjunto3,'') As Adjunto3,
                        ISNULL(Adjunto4,'') As Adjunto4,
                        ISNULL(Adjunto5,'') As Adjunto5,
                        ISNULL(Adjunto6,'') As Adjunto6,
                        ISNULL(Adjunto7,'') As Adjunto7,
                        ISNULL(Adjunto8,'') As Adjunto8,
                        ISNULL(Adjunto9,'') As Adjunto9,
                        ISNULL(Adjunto10,'') As Adjunto10,
                        ISNULL(Adjunto11,'') As Adjunto11,
                        ISNULL(Adjunto12,'') As Adjunto12,
                        ISNULL(Adjunto13,'') As Adjunto13,
                        ISNULL(Adjunto14,'') As Adjunto14,
                        ISNULL(Adjunto15,'') As Adjunto15,
                        ISNULL(Adjunto16,'') As Adjunto16,
                        ISNULL(Adjunto17,'') As Adjunto17,
                        ISNULL(Adjunto18,'') As Adjunto18,
                        ISNULL(Adjunto19,'') As Adjunto19,
                        ISNULL(Adjunto20,'') As Adjunto20,
                        ISNULL(Adjunto21,'') As Adjunto21,
                        ISNULL(Adjunto22,'') As Adjunto22,
                        ISNULL(Adjunto23,'') As Adjunto23,
                        ISNULL(Adjunto24,'') As Adjunto24,
                        ISNULL(Adjunto25,'') As Adjunto25,
                        EmailEmisor,
                        SmtpEmisor,
                        PassEmisor,
                        MRH_Interno,
                        FirmaEmisor,MRH_OrigenNotificacion,
                        TelegramID
                        FROM Aurora.dbo.MRH_Notificaciones WHERE FechaConfirmadaEnvio IS NULL AND EnviaEmail = -1 AND ErrorEnvio = 0"
            sqlcomando = New SqlCommand()
            sqlcomando.CommandText = consulta
            sqlcomando.CommandType = CommandType.Text
            sqlcomando.Connection = conexion

            respuesta = sqlcomando.ExecuteReader

            While respuesta.Read
                MovPosicion = respuesta.Item("MovPosicion")
                Email = respuesta.Item("Email")
                Asunto = respuesta.Item("Asunto")
                Mensaje = respuesta.Item("Mensaje")
                Adjunto = respuesta.Item("Adjunto")
                Adjunto1 = respuesta.Item("Adjunto1")
                Adjunto2 = respuesta.Item("Adjunto2")
                Adjunto3 = respuesta.Item("Adjunto3")
                Adjunto4 = respuesta.Item("Adjunto4")
                Adjunto5 = respuesta.Item("Adjunto5")
                Adjunto6 = respuesta.Item("Adjunto6")
                Adjunto7 = respuesta.Item("Adjunto7")
                Adjunto8 = respuesta.Item("Adjunto8")
                Adjunto9 = respuesta.Item("Adjunto9")
                Adjunto10 = respuesta.Item("Adjunto10")
                Adjunto11 = respuesta.Item("Adjunto11")
                Adjunto12 = respuesta.Item("Adjunto12")
                Adjunto13 = respuesta.Item("Adjunto13")
                Adjunto14 = respuesta.Item("Adjunto14")
                Adjunto15 = respuesta.Item("Adjunto15")
                Adjunto16 = respuesta.Item("Adjunto16")
                Adjunto17 = respuesta.Item("Adjunto17")
                Adjunto18 = respuesta.Item("Adjunto18")
                Adjunto19 = respuesta.Item("Adjunto19")
                Adjunto20 = respuesta.Item("Adjunto20")
                Adjunto21 = respuesta.Item("Adjunto21")
                Adjunto22 = respuesta.Item("Adjunto22")
                Adjunto23 = respuesta.Item("Adjunto23")
                Adjunto24 = respuesta.Item("Adjunto24")
                Adjunto25 = respuesta.Item("Adjunto25")
                MRH_interno = respuesta.Item("MRH_interno")
                EmailEmisor = respuesta.Item("EmailEmisor")
                PassEmisor = respuesta.Item("PassEmisor")
                SmtpEmisor = respuesta.Item("SmtpEmisor")
                FirmaEmisor = respuesta.Item("FirmaEmisor")
                TelegramID = respuesta.Item("TelegramID")
                MRH_OrigenNotificacion = respuesta.Item("MRH_OrigenNotificacion")
                Try
                    If MRH_interno = -1 Then
                        cabeceraMensaje = "<html><body><table width='90%' cellpadding='2' cellspacing='2' border='0'><tr><td><img src=""cid:Pic1""></td><td><h1>SISTEMA DE NOTIFICACIONES</h1></td></tr></table>"
                        cuerpoMensaje = Mensaje
                        pieMensaje = "</body></html>"
                        htmlBody = cabeceraMensaje + "<br><br>" + cuerpoMensaje + pieMensaje
                        avHtml = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, System.Net.Mime.MediaTypeNames.Text.Html)
                        pic1 = New LinkedResource("C:\MRH\Servicios\LogoBoton.png", System.Net.Mime.MediaTypeNames.Image.Jpeg)
                        pic1.ContentId = "Pic1"
                        avHtml.LinkedResources.Add(pic1)
                        textBody = "You must use an e-mail client that supports HTML messages"
                        m = New MailMessage
                        m.AlternateViews.Add(avHtml)
                        m.From = New MailAddress("aurorabot@auroracorp.es")
                        m.To.Add(Email)
                        m.Subject = Asunto
                        If Adjunto <> "" Then
                            Dim archi As New Attachment(Adjunto)
                            m.Attachments.Add(archi)
                        End If
                        If Adjunto1 <> "" Then
                            Dim archi1 As New Attachment(Adjunto1)
                            m.Attachments.Add(archi1)
                        End If
                        If Adjunto2 <> "" Then
                            Dim archi2 As New Attachment(Adjunto2)
                            m.Attachments.Add(archi2)
                        End If
                        If Adjunto3 <> "" Then
                            Dim archi3 As New Attachment(Adjunto3)
                            m.Attachments.Add(archi3)
                        End If
                        If Adjunto4 <> "" Then
                            Dim archi4 As New Attachment(Adjunto4)
                            m.Attachments.Add(archi4)
                        End If
                        If Adjunto5 <> "" Then
                            Dim archi5 As New Attachment(Adjunto5)
                            m.Attachments.Add(archi5)
                        End If
                        If Adjunto6 <> "" Then
                            Dim archi6 As New Attachment(Adjunto6)
                            m.Attachments.Add(archi6)
                        End If
                        If Adjunto7 <> "" Then
                            Dim archi7 As New Attachment(Adjunto7)
                            m.Attachments.Add(archi7)
                        End If
                        If Adjunto8 <> "" Then
                            Dim archi8 As New Attachment(Adjunto8)
                            m.Attachments.Add(archi8)
                        End If
                        If Adjunto9 <> "" Then
                            Dim archi9 As New Attachment(Adjunto9)
                            m.Attachments.Add(archi9)
                        End If
                        If Adjunto10 <> "" Then
                            Dim archi10 As New Attachment(Adjunto10)
                            m.Attachments.Add(archi10)
                        End If
                        If Adjunto11 <> "" Then
                            Dim archi11 As New Attachment(Adjunto11)
                            m.Attachments.Add(archi11)
                        End If
                        If Adjunto12 <> "" Then
                            Dim archi12 As New Attachment(Adjunto12)
                            m.Attachments.Add(archi12)
                        End If
                        If Adjunto13 <> "" Then
                            Dim archi13 As New Attachment(Adjunto13)
                            m.Attachments.Add(archi13)
                        End If
                        If Adjunto14 <> "" Then
                            Dim archi14 As New Attachment(Adjunto14)
                            m.Attachments.Add(archi14)
                        End If
                        If Adjunto15 <> "" Then
                            Dim archi15 As New Attachment(Adjunto15)
                            m.Attachments.Add(archi15)
                        End If
                        If Adjunto16 <> "" Then
                            Dim archi16 As New Attachment(Adjunto16)
                            m.Attachments.Add(archi16)
                        End If
                        If Adjunto17 <> "" Then
                            Dim archi17 As New Attachment(Adjunto17)
                            m.Attachments.Add(archi17)
                        End If
                        If Adjunto18 <> "" Then
                            Dim archi18 As New Attachment(Adjunto18)
                            m.Attachments.Add(archi18)
                        End If
                        If Adjunto19 <> "" Then
                            Dim archi19 As New Attachment(Adjunto19)
                            m.Attachments.Add(archi19)
                        End If
                        If Adjunto20 <> "" Then
                            Dim archi20 As New Attachment(Adjunto20)
                            m.Attachments.Add(archi20)
                        End If
                        If Adjunto21 <> "" Then
                            Dim archi21 As New Attachment(Adjunto21)
                            m.Attachments.Add(archi21)
                        End If
                        If Adjunto22 <> "" Then
                            Dim archi22 As New Attachment(Adjunto22)
                            m.Attachments.Add(archi22)
                        End If
                        If Adjunto23 <> "" Then
                            Dim archi23 As New Attachment(Adjunto23)
                            m.Attachments.Add(archi23)
                        End If
                        If Adjunto24 <> "" Then
                            Dim archi24 As New Attachment(Adjunto24)
                            m.Attachments.Add(archi24)
                        End If
                        If Adjunto25 <> "" Then
                            Dim archi25 As New Attachment(Adjunto25)
                            m.Attachments.Add(archi25)
                        End If
                        client = New SmtpClient("aurorabot.auroracorp.es")
                        client.Credentials = New Net.NetworkCredential("news_auroraco44", "Aur0r4Crg2022")
                        client.Send(m)
                        Funciones.UpdateNotificacion(MovPosicion)
                    End If
                    If MRH_interno = 0 Then
                        cabeceraMensaje = "<html><body>"
                        cuerpoMensaje = Mensaje
                        pieMensaje = FirmaEmisor + "</body></html>"
                        htmlBody = cabeceraMensaje + "<br><br>" + cuerpoMensaje + "<br><br>" + pieMensaje
                        avHtml = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, System.Net.Mime.MediaTypeNames.Text.Html)

                        textBody = "You must use an e-mail client that supports HTML messages"
                        m = New MailMessage
                        m.AlternateViews.Add(avHtml)
                        m.From = New MailAddress(EmailEmisor)
                        m.To.Add(Email)
                        m.Subject = Asunto

                        client = New SmtpClient(SmtpEmisor)
                        If MRH_OrigenNotificacion = "SOPORTE_IT" Then
                            client.Credentials = New Net.NetworkCredential("news_aurorain63", PassEmisor)
                        Else
                            client.Credentials = New Net.NetworkCredential(EmailEmisor, PassEmisor)
                        End If
                        client.Send(m)

                        Funciones.UpdateNotificacion(MovPosicion)
                    End If
                Catch
                    Funciones.UpdateNotificacionError(MovPosicion)



                End Try
            End While
            respuesta.Close()
            conexion.Close()
            conexion.Dispose()
        Catch ex As Exception
            '    
        End Try
    End Sub




    'Funcion Para hacer los envíos de los mensajes de telegram de forma independiente

    Private Async Sub TimerLoop_Tick2(sender As Object, e As EventArgs)
        Dim conexion As SqlConnection
        Dim sqlcomando As SqlCommand
        Dim consulta As String
        Dim respuesta As IDataReader
        Dim MovPosicion As Guid
        Dim Email As String = ""
        Dim MRH_OrigenNotificacion As String = ""
        Dim Asunto As String = ""
        Dim Mensaje As String = ""
        Dim Adjunto As String = ""
        Dim Adjunto1 As String = ""
        Dim Adjunto2 As String = ""
        Dim Adjunto3 As String = ""
        Dim Adjunto4 As String = ""
        Dim Adjunto5 As String = ""
        Dim Adjunto6 As String = ""
        Dim Adjunto7 As String = ""
        Dim Adjunto8 As String = ""
        Dim Adjunto9 As String = ""
        Dim Adjunto10 As String = ""
        Dim Adjunto11 As String = ""
        Dim Adjunto12 As String = ""
        Dim Adjunto13 As String = ""
        Dim Adjunto14 As String = ""
        Dim Adjunto15 As String = ""
        Dim Adjunto16 As String = ""
        Dim Adjunto17 As String = ""
        Dim Adjunto18 As String = ""
        Dim Adjunto19 As String = ""
        Dim Adjunto20 As String = ""
        Dim Adjunto21 As String = ""
        Dim Adjunto22 As String = ""
        Dim Adjunto23 As String = ""
        Dim Adjunto24 As String = ""
        Dim Adjunto25 As String = ""
        Dim MRH_interno As Integer = 0
        Dim EmailEmisor As String = ""
        Dim PassEmisor As String = ""
        Dim SmtpEmisor As String = ""
        Dim FirmaEmisor As String = ""
        Dim cabeceraMensaje As String = ""
        Dim cuerpoMensaje As String = ""
        Dim pieMensaje As String = ""
        Dim htmlBody As String = ""
        Dim avHtml As AlternateView
        Dim pic1 As LinkedResource
        Dim textBody As String = ""
        Dim TelegramID As String = ""
        Dim m As MailMessage
        Dim client As SmtpClient

        'Funciones.EnviaNotificaciones()
        Try
            conexion = New SqlConnection(Funciones.CadenaConexionSage)
            conexion.Open()
            consulta = "SELECT TOP 1 MovPosicion,Email,Asunto,Mensaje,
                        ISNULL(Adjunto,'') As Adjunto,
                        ISNULL(Adjunto1,'') As Adjunto1,
                        ISNULL(Adjunto2,'') As Adjunto2,
                        ISNULL(Adjunto3,'') As Adjunto3,
                        ISNULL(Adjunto4,'') As Adjunto4,
                        ISNULL(Adjunto5,'') As Adjunto5,
                        ISNULL(Adjunto6,'') As Adjunto6,
                        ISNULL(Adjunto7,'') As Adjunto7,
                        ISNULL(Adjunto8,'') As Adjunto8,
                        ISNULL(Adjunto9,'') As Adjunto9,
                        ISNULL(Adjunto10,'') As Adjunto10,
                        ISNULL(Adjunto11,'') As Adjunto11,
                        ISNULL(Adjunto12,'') As Adjunto12,
                        ISNULL(Adjunto13,'') As Adjunto13,
                        ISNULL(Adjunto14,'') As Adjunto14,
                        ISNULL(Adjunto15,'') As Adjunto15,
                        ISNULL(Adjunto16,'') As Adjunto16,
                        ISNULL(Adjunto17,'') As Adjunto17,
                        ISNULL(Adjunto18,'') As Adjunto18,
                        ISNULL(Adjunto19,'') As Adjunto19,
                        ISNULL(Adjunto20,'') As Adjunto20,
                        ISNULL(Adjunto21,'') As Adjunto21,
                        ISNULL(Adjunto22,'') As Adjunto22,
                        ISNULL(Adjunto23,'') As Adjunto23,
                        ISNULL(Adjunto24,'') As Adjunto24,
                        ISNULL(Adjunto25,'') As Adjunto25,
                        EmailEmisor,
                        SmtpEmisor,
                        PassEmisor,
                        MRH_Interno,
                        FirmaEmisor,MRH_OrigenNotificacion,
                        TelegramID
                        FROM Aurora.dbo.MRH_Notificaciones WHERE FechaConfirmadaEnvioT IS NULL AND EnviaApp = -1 AND ErrorEnvio = 0"
            sqlcomando = New SqlCommand()
            sqlcomando.CommandText = consulta
            sqlcomando.CommandType = CommandType.Text
            sqlcomando.Connection = conexion

            respuesta = sqlcomando.ExecuteReader

            While respuesta.Read
                MovPosicion = respuesta.Item("MovPosicion")
                Email = respuesta.Item("Email")
                Asunto = respuesta.Item("Asunto")
                Mensaje = respuesta.Item("Mensaje")
                Adjunto = respuesta.Item("Adjunto")
                Adjunto1 = respuesta.Item("Adjunto1")
                Adjunto2 = respuesta.Item("Adjunto2")
                Adjunto3 = respuesta.Item("Adjunto3")
                Adjunto4 = respuesta.Item("Adjunto4")
                Adjunto5 = respuesta.Item("Adjunto5")
                Adjunto6 = respuesta.Item("Adjunto6")
                Adjunto7 = respuesta.Item("Adjunto7")
                Adjunto8 = respuesta.Item("Adjunto8")
                Adjunto9 = respuesta.Item("Adjunto9")
                Adjunto10 = respuesta.Item("Adjunto10")
                Adjunto11 = respuesta.Item("Adjunto11")
                Adjunto12 = respuesta.Item("Adjunto12")
                Adjunto13 = respuesta.Item("Adjunto13")
                Adjunto14 = respuesta.Item("Adjunto14")
                Adjunto15 = respuesta.Item("Adjunto15")
                Adjunto16 = respuesta.Item("Adjunto16")
                Adjunto17 = respuesta.Item("Adjunto17")
                Adjunto18 = respuesta.Item("Adjunto18")
                Adjunto19 = respuesta.Item("Adjunto19")
                Adjunto20 = respuesta.Item("Adjunto20")
                Adjunto21 = respuesta.Item("Adjunto21")
                Adjunto22 = respuesta.Item("Adjunto22")
                Adjunto23 = respuesta.Item("Adjunto23")
                Adjunto24 = respuesta.Item("Adjunto24")
                Adjunto25 = respuesta.Item("Adjunto25")
                MRH_interno = respuesta.Item("MRH_interno")
                EmailEmisor = respuesta.Item("EmailEmisor")
                PassEmisor = respuesta.Item("PassEmisor")
                SmtpEmisor = respuesta.Item("SmtpEmisor")
                FirmaEmisor = respuesta.Item("FirmaEmisor")
                TelegramID = respuesta.Item("TelegramID")
                MRH_OrigenNotificacion = respuesta.Item("MRH_OrigenNotificacion")
                Try

                    'Despues de enviar la notificación por correo, enviamos el mensaje por telegram. 

                    Await Bot_client.SendTextMessageAsync(TelegramID, Mensaje)


                    Funciones.UpdateNotificacionT(MovPosicion)

                Catch
                    Funciones.UpdateNotificacionError(MovPosicion)



                End Try
            End While
            respuesta.Close()
            conexion.Close()
            conexion.Dispose()
        Catch ex As Exception
            '    
        End Try
    End Sub


    'Acabo funcion




End Class
