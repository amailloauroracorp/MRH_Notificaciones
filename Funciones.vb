Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Telegram.Bot
Imports System.IO
Imports Telegram.Bot.Types

Public Class Funciones
    Shared Bot_client As TelegramBotClient = New TelegramBotClient("5770884977:AAGDI5TJ5vn_e_DYtumrfqrrY6iFX9dClWg")
    Public Shared CadenaConexionSage As String = "Data Source=SAGE200\SAGE;User ID=MRH;Password=Aurora_2019_1+#;"
    Public Shared Function EnviaNotificaciones()
        Dim resultado As Integer = 0





        Return resultado
    End Function

    Public Shared Function UpdateNotificacion(ByVal MovPosicion As Guid)
        Dim resultado As Integer
        resultado = 0
        Dim conexion As SqlConnection
        Dim sqlcomando As SqlCommand
        Dim consulta As String
        conexion = New SqlConnection(CadenaConexionSage)
        conexion.Open()
        consulta = "UPDATE AURORA.DBO.MRH_Notificaciones SET FechaConfirmadaEnvio = GETDATE() WHERE MovPosicion = '" + MovPosicion.ToString + "'"
        sqlcomando = New SqlCommand()
        sqlcomando.CommandText = consulta
        sqlcomando.CommandType = CommandType.Text
        sqlcomando.Connection = conexion
        sqlcomando.ExecuteNonQuery()
        conexion.Close()
        conexion.Dispose()

        Return resultado
    End Function
    Public Shared Function UpdateNotificacionT(ByVal MovPosicion As Guid)
        Dim resultado As Integer
        resultado = 0
        Dim conexion As SqlConnection
        Dim sqlcomando As SqlCommand
        Dim consulta As String
        conexion = New SqlConnection(CadenaConexionSage)
        conexion.Open()
        consulta = "UPDATE AURORA.DBO.MRH_Notificaciones SET FechaConfirmadaEnvioT = GETDATE() WHERE MovPosicion = '" + MovPosicion.ToString + "'"
        sqlcomando = New SqlCommand()
        sqlcomando.CommandText = consulta
        sqlcomando.CommandType = CommandType.Text
        sqlcomando.Connection = conexion
        sqlcomando.ExecuteNonQuery()
        conexion.Close()
        conexion.Dispose()

        Return resultado
    End Function

    Public Shared Function UpdateNotificacionError(ByVal MovPosicion As Guid)
        Dim resultado As Integer
        resultado = 0
        Dim conexion As SqlConnection
        Dim sqlcomando As SqlCommand
        Dim consulta As String
        conexion = New SqlConnection(CadenaConexionSage)
        conexion.Open()
        consulta = "UPDATE AURORA.DBO.MRH_Notificaciones SET ErrorEnvio = -1 WHERE MovPosicion = '" + MovPosicion.ToString + "'"
        sqlcomando = New SqlCommand()
        sqlcomando.CommandText = consulta
        sqlcomando.CommandType = CommandType.Text
        sqlcomando.Connection = conexion
        sqlcomando.ExecuteNonQuery()
        conexion.Close()
        conexion.Dispose()

        Return resultado
    End Function

    Public Shared Async Function EnviaAdjuntoTelegram(TelegramID As String, RutaArchivo As String) As Task


        'Si el archivo definido en los campos de AdjuntoX de MRH.Notificaciones Existe, lo cargo en el FileStream y luego lo mando 
        If System.IO.File.Exists(RutaArchivo) Then

            Dim InputFileStream As New InputFileStream(New FileStream(RutaArchivo, FileMode.Open), Path.GetFileName(RutaArchivo))
            Await Bot_client.SendDocumentAsync(TelegramID, InputFileStream)

        End If



    End Function



End Class




