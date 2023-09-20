Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class Funciones
    Public Shared CadenaConexionSage As String = "Data Source=SAGE200\SAGE;User ID=logic;Password=Sage2009+;"
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
End Class
