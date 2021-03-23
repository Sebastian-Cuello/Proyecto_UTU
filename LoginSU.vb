Imports System.Data.Odbc

Public Class LoginSU
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet
    Dim read As OdbcDataReader
    Dim eleccion As Integer = 0 'La eleccion de perfil 

    Dim perfil As String
    Dim anda As Boolean = False

    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=" & txtUsuario.Text & ";PWD=" & txtContraseña.Text
            cx.Open()
            cm.Connection = cx
            anda = True
        Catch
            MsgBox("Nº Error 0: No hay conexion con la base de datos.")
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click 'Admin
        txtContraseña.Show()
        txtUsuario.Show()
        labelUsu.Show()
        labelCon.Show()
        labelCedula.Hide()
        txtInvi.Hide()

        txtContraseña.Text = ""
        txtUsuario.Text = ""
        eleccion = 1

        lblPerfil.Text = "Bienvenido Administrador"
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click 'Vendedor
        txtContraseña.Show()
        txtUsuario.Show()
        labelUsu.Show()
        labelCon.Show()
        labelCedula.Hide()
        txtInvi.Hide()

        txtContraseña.Text = ""
        txtUsuario.Text = ""
        eleccion = 2

        lblPerfil.Text = "Bienvenido Recepcionista"
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click 'Mecanico
        txtContraseña.Show()
        txtUsuario.Show()
        labelUsu.Show()
        labelCon.Show()
        labelCedula.Hide()
        txtInvi.Hide()

        txtContraseña.Text = ""
        txtUsuario.Text = ""
        eleccion = 3

        lblPerfil.Text = "Bienvenido Mecanico"
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click 'Invitado
        eleccion = 4
        lblPerfil.Text = "Bienvenido Invitado"

        txtContraseña.Text = ""
        txtUsuario.Text = ""
        txtContraseña.Hide()
        txtUsuario.Hide()
        labelUsu.Hide()
        labelCon.Hide()
        labelCedula.Show()
        txtInvi.Show()
    End Sub

    Public user As String
    Public contraseña As String

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click 'Boton login
        login()

    End Sub

    '-----------------------------------HOVER MOUSE------------------------------------------------------------------------------
    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover 'Cuando pasa el mouse por encima del imagen admin
        PictureBox1.BackColor = Color.White
    End Sub

    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave 'Cuando quita el mouse sobre la imagen admin
        PictureBox1.BackColor = Color.FromArgb(255, 192, 128)
    End Sub

    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover 'Cuando pasa el mouse por encima del imagen vendedor
        PictureBox2.BackColor = Color.White
    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave 'Cuando quita el mouse sobre la imagen vendedor
        PictureBox2.BackColor = Color.FromArgb(255, 192, 128)
    End Sub

    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs) Handles PictureBox3.MouseHover 'Cuando pasa el mouse por encima del imagen mecanico
        PictureBox3.BackColor = Color.White
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave 'Cuando quita el mouse sobre la imagen mecanico
        PictureBox3.BackColor = Color.FromArgb(255, 192, 128)
    End Sub

    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover 'Cuando pasa el mouse por encima del imagen invitado
        PictureBox4.BackColor = Color.White
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave 'Cuando quita el mouse sobre la imagen invitado
        PictureBox4.BackColor = Color.FromArgb(255, 192, 128)
    End Sub

    Private Sub LoginSU_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtContraseña.UseSystemPasswordChar = True
        Me.MaximizeBox = False
    End Sub

    Private Sub txtContraseña_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtContraseña.KeyPress, txtInvi.KeyPress
        If Asc(e.KeyChar) = 13 Then 'cuando le da enter
            login()
        End If
    End Sub

    Private Sub login()
        If Not eleccion = 4 Then

            conexion()
            If anda = True Then
                'obtener perfil de ususario y contraseña, luego comparar dependiendo de la eleccion del perfil
                Dim continua As Boolean = False
                user = txtUsuario.Text
                contraseña = txtContraseña.Text

                If Not eleccion = 4 Then
                    Try
                        cm.CommandText = "SELECT defrole FROM sysusers where username  = " & Chr(34) & txtUsuario.Text & Chr(34)
                        read = cm.ExecuteReader
                        read.Read()
                        perfil = read.Item(0).ToString.Trim
                        continua = True
                    Catch
                        MsgBox("Nº Error 1: Usuario o Contraseña equivocado.")
                    End Try
                    cx.Close()
                End If

                If eleccion = 4 Then
                    continua = True
                End If

                If continua = True Then
                    If eleccion = 0 Then
                        MsgBox("Elija su perfil.")
                    ElseIf eleccion = 1 Then
                        If perfil = "administrador" Then
                            Me.Hide()
                            ventanaAdmin.Show()
                        Else
                            MsgBox("Nº Error 2: Usuario o perfil incorrecto.")
                        End If
                    ElseIf eleccion = 2 Then
                        If perfil = "recepcionista" Then
                            Me.Hide()
                            VentanaRecepcion.Show()
                        Else
                            MsgBox("Nº Error 3: Usuario o perfil incorrecto.")
                        End If
                    ElseIf eleccion = 3 Then
                        If perfil = "mecanico" Then
                            Me.Hide()
                            ventanaMecanico.Show()
                        Else
                            MsgBox("Nº Error 4: Usuario o perfil incorrecto.")
                        End If
                    End If
                End If
            End If
        End If

        If eleccion = 4 Then
            '--------------algoritmo comprobacion de CI-------------
            Dim cedula As Char() = txtInvi.Text
            Dim numeroci, total As Integer

            Try
                numeroci = Int32.Parse(cedula)
                For i = 0 To 6 Step 1
                    numeroci = Int32.Parse(cedula(i))

                    Select Case i
                        Case 0
                            numeroci = numeroci * 2
                        Case 1
                            numeroci = numeroci * 9
                        Case 2
                            numeroci = numeroci * 8
                        Case 3
                            numeroci = numeroci * 7
                        Case 4
                            numeroci = numeroci * 6
                        Case 5
                            numeroci = numeroci * 3
                        Case 6
                            numeroci = numeroci * 4
                    End Select
                    total = total + numeroci
                Next

                If total Mod 10 + Int32.Parse(cedula(7)) = 10 Then
                    Me.Hide()
                    Invitado.Show()
                Else
                    MsgBox("Nº Error 5: Ingrese una Cedula Valida.")
                End If
            Catch
                MsgBox("Nº Error 6: Ingrese una Cedula Valida.")
            End Try
            '-------------------
        End If
    End Sub
End Class