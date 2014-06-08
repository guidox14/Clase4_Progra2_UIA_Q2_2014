Imports System.Data.SqlClient

Public Class Login

    ''' <summary>
    ''' Controla el evento de clic del botón
    ''' </summary>
    ''' <param name="sender">Objeto que envía el evento</param>
    ''' <param name="e">Argumentos del evento</param>
    ''' <remarks></remarks>
    Private Sub Login_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Login_Button.Click
        Dim nombreUsuario As String
        Dim password As String

        nombreUsuario = NombreUsuario_TextBox.Text
        password = Password_TextBox.Text

        If (IsNumeric(password)) Then
            'Llama al método que hace la conexión a la base de datos
            If (LoginBD(nombreUsuario, password)) Then
                Dim loginCorrecto As String = "El usuario " + nombreUsuario + " ha ingresado al programa."
                MsgBox(loginCorrecto, MsgBoxStyle.Information, "Acceso Correcto")
            Else
                'Ejecuta esta línea en caso que el método devuelva falso, lo cual pasa si el select no devolvió ningún dato.
                Dim noLogin = "El usuario o la contraseña están incorrectos"
                MsgBox(noLogin, MsgBoxStyle.Critical, "Error en el login")
            End If
        Else
            'Ejecuta las líneas que siguen si el password no viene en formato númerico
            Dim mensajeError = "El password no viene en formato numerico"
            MsgBox(mensajeError, MsgBoxStyle.Exclamation, "Error en el login")
        End If

    End Sub

    ''' <summary>
    ''' Se conecta a base de datos y ejecuta un select
    ''' </summary>
    ''' <param name="nombre">Nombre del usuario proveniente de la interfaz gráfica</param>
    ''' <param name="pass">Password del usuario, proveniente de la interfaz gráfica</param>
    ''' <returns>True si el select devuelve campos, false si el select no devuelve nada</returns>
    ''' <remarks></remarks>
    Private Function LoginBD(ByVal nombre As String, ByVal pass As String) As Boolean
        'Crea la variable 'connectionBuilder', la cual se convierte en una instancia de la clase SqlConnectionStringBuilder
        Dim connectionBuilder As New SqlConnectionStringBuilder
        connectionBuilder.DataSource = "ADRIAN-PC\SQLEXPRESS" 'Establece el data source al cual se va a conectar
        connectionBuilder.InitialCatalog = "Practica_Clase4" 'Establece el nombre de la base de datos que va a utilizar
        connectionBuilder.IntegratedSecurity = True 'Establece si utiliza seguridad integrada, esto es, la seguridad con que se hace la conexión a la base de datos

        'Crea una variable tipo SqlConnection, para poderse conectar hacia la base de datos
        Dim conexion As New SqlConnection(connectionBuilder.ConnectionString)
        conexion.Open() 'Abre la conexión

        'Crea una variable de tipo SqlCommand
        Dim sqlCommand As SqlCommand = conexion.CreateCommand
        sqlCommand.Connection = conexion

        'Asigna el texto del select a la variable command de tipo String
        Dim command As String = "SELECT NombreLogin" & _
                            ", Pass" & _
                            " FROM Usuarios WHERE NombreLogin = '" + nombre + "' and Pass = '" + pass + "';"
        sqlCommand.CommandText = command
        Dim sqlReader As SqlDataReader

        'Ejecuta el ExecuteReader y trae de vuelta los datos leídos de la base de datos.
        sqlReader = sqlCommand.ExecuteReader()

        'Pregunta si sqlReader tiene un elemento que leer. Sí lo tiene, entonces la función devuelve 'true'. Caso contrario devuelve 'false'
        If (sqlReader.Read()) Then
            conexion.Close() 'Cierra la conexión
            Return True
        Else
            conexion.Close()
            Return False
        End If

    End Function


End Class
