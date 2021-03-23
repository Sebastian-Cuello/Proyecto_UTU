Imports System.Data.Odbc
Public Class VentanaRecepcion
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet
    Dim cm2 As New Odbc.OdbcCommand
    Dim da2 As New Odbc.OdbcDataAdapter
    Dim ds2 As New Data.DataSet
    Dim read As OdbcDataReader
    Dim abierto As Boolean = False 'Si hay un formulario abierto será verdadero
    Dim caso1 As Boolean = False
    Dim caso2 As Boolean = False
    Dim caso3 As Boolean = False
    Dim caso4 As Boolean = False
    Dim opcion As Integer 'opcion del menu para desplegar submenus
    Private Sub pictureCerrarSesion_Click(sender As Object, e As EventArgs) Handles pictureCerrarSesion.Click 'flecha cerrar sesion
        Me.Close()
        LoginSU.Show()
    End Sub


    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & application.startupPath & "\odbcCharm.dsn;UID=" & LoginSU.user & ";PWD=" & LoginSU.contraseña
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 45: No hay conexion con la base de datos.")
        End Try
    End Sub

    Public Sub historial(comando As String) 'mantiene la tabla historial al dia
        'armar sentencia de insercion
        Dim host As String = System.Net.Dns.GetHostName()
        Dim LocalHostaddress As String = System.Net.Dns.GetHostByName(host).AddressList(0).ToString()
        Dim hora As String = DateAndTime.Now.ToString("yyyy-MM-dd hh:mm")

        Dim arra As Char() = comando

        For i = 0 To arra.Length() - 1
            If arra(i) = Chr(34) Then
                arra(i) = "'"
            End If
        Next

        conexion()
        Try
            cm.CommandText = "insert into historial values (0," & Chr(34) & LocalHostaddress & Chr(34) & "," & Chr(34) & LoginSU.user & Chr(34) & "," & Chr(34) &
            arra & Chr(34) & "," & Chr(34) & hora & Chr(34) & ");"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 46: Error al mantener historial.")
        End Try
        cx.Close()
    End Sub

    Private Sub ventanaRecepcion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = False

        conexion()
        Try 'carga combobox sucursales vehiculo
            cm.CommandText = "SELECT nombre_sucursal FROM sucursal"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds)
            ComboBoxSucursalV.DataSource = ds.Tables(0)
            ComboBoxSucursalV.DisplayMember = ds.Tables(0).Columns(0).Caption.ToString
            ComboBoxSucursal.DataSource = ds.Tables(0)
            ComboBoxSucursal.DisplayMember = ds.Tables(0).Columns(0).Caption.ToString
        Catch
            MsgBox("N°Error 47: Error al cargar sucursales.")
        End Try
        cx.Close()

        conexion()
        Try 'carga combobox categoria reserva 
            cm2.Connection = cx
            cm2.CommandText = "SELECT nombre_categoria FROM categoria where estado = 'Activo'"
            cm2.ExecuteNonQuery()
            da2.SelectCommand = cm2
            da2.Fill(ds2)
            ComboBoxCategoriaVehiculo.DataSource = ds2.Tables(0)
            ComboBoxCategoriaVehiculo.DisplayMember = ds2.Tables(0).Columns(0).Caption.ToString
            ComboBoxCategoria.DataSource = ds2.Tables(0)
            ComboBoxCategoria.DisplayMember = ds2.Tables(0).Columns(0).Caption.ToString
        Catch
            MsgBox("N°Error 48: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategoriasListado.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 49: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM vehiculo where disponibilidad = 'Disponible'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 50: Error al cargar vehiculos.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM cliente"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Cliente")
            DataGridListadoCliente.DataSource = ds.Tables("Cliente")
        Catch
            MsgBox("N°Error 51: Error al cargar clientes.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM reserva"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Reserva")
            DataGridListadoReserva.DataSource = ds.Tables("Reserva")
        Catch
            MsgBox("N°Error 52: Error al cargar reservas.")
        End Try
        cx.Close()

    End Sub

    '------------------------------Botones En Imagenes---------------------------------------------------------
    Private Sub botonesOpciones() 'Abre el boton correcto de la opcion y esconde los otros
        Select Case opcion
            Case 1
                If caso1 = False Then
                    picReservaGestion.Show()
                    picReservaListado.Show()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso1 = True
                Else
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    caso1 = False
                End If
            Case 2
                If caso2 = False Then
                    picClienteListado.Show()
                    picClienteGestion.Show()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso2 = True
                Else
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    caso2 = False
                End If
            Case 3
                If caso3 = False Then
                    picVehiculoGestion.Show()
                    picVehiculoListado.Show()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso3 = True
                Else
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    caso3 = False
                End If
            Case 4
                If caso4 = False Then
                    picCategoriaGestion.Show()
                    picCategoriaListado.Show()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    caso4 = True
                Else
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso4 = False
                End If
        End Select
    End Sub

    Private Sub btnReserva_Click(sender As Object, e As EventArgs) Handles btnReserva.Click        'click boton reserva
        opcion = 1
        botonesOpciones()
    End Sub

    Private Sub picReservaGestion_Click(sender As Object, e As EventArgs) Handles picReservaGestion.Click        'Abrir form gestion reserva
        picReservaGestion.Hide()
        picReservaListado.Hide()
        If abierto = False Then
            abierto = True
            panelReserva.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelReserva.Location.X = 159
                panelReserva.Location = New Point(panelReserva.Location.X + 1, 0)
            Loop

            Do Until panelReserva.Location.X = 169
                panelReserva.Location = New Point(panelReserva.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop

        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click       'Cerrar Formulario gestion Reserva
        abierto = False
        charmLogo.Visible = False
        panelReserva.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelReserva.Location.X = -576
            panelReserva.Location = New Point(panelReserva.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelReserva.Location.X = -586
            panelReserva.Location = New Point(panelReserva.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picReservaListado_Click(sender As Object, e As EventArgs) Handles picReservaListado.Click        'Abrir listado Reserva
        picReservaGestion.Hide()
        picReservaListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoReserva.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoReserva.Location.X = 159
                panelListadoReserva.Location = New Point(panelListadoReserva.Location.X + 1, 0)
            Loop

            Do Until panelListadoReserva.Location.X = 169
                panelListadoReserva.Location = New Point(panelListadoReserva.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop

        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click 'Cerrar listado Reserva
        abierto = False
        charmLogo.Visible = False
        panelListadoReserva.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoReserva.Location.X = -576
            panelListadoReserva.Location = New Point(panelListadoReserva.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelListadoReserva.Location.X = -586
            panelListadoReserva.Location = New Point(panelListadoReserva.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click        'click boton cliente
        opcion = 2
        botonesOpciones()
    End Sub

    Private Sub picClienteGestion_Click(sender As Object, e As EventArgs) Handles picClienteGestion.Click        'Abrir form cliente
        picClienteListado.Hide()
        picClienteGestion.Hide()
        If abierto = False Then
            abierto = True
            panelCliente.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelCliente.Location.X = 159
                panelCliente.Location = New Point(panelCliente.Location.X + 1, 0)
            Loop

            Do Until panelCliente.Location.X = 169
                panelCliente.Location = New Point(panelCliente.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop

        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click 'Cerrar form cliente
        abierto = False
        charmLogo.Visible = False
        panelCliente.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelCliente.Location.X = -576
            panelCliente.Location = New Point(panelCliente.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelCliente.Location.X = -586
            panelCliente.Location = New Point(panelCliente.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picClienteListado_Click(sender As Object, e As EventArgs) Handles picClienteListado.Click        'abrir Form listado cliente
        picClienteListado.Hide()
        picClienteGestion.Hide()
        If abierto = False Then
            abierto = True
            panelListadoCliente.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoCliente.Location.X = 159
                panelListadoCliente.Location = New Point(panelListadoCliente.Location.X + 1, 0)
            Loop

            Do Until panelListadoCliente.Location.X = 169
                panelListadoCliente.Location = New Point(panelListadoCliente.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click 'cerrar form listado cliente
        abierto = False
        charmLogo.Visible = False
        panelListadoCliente.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoCliente.Location.X = -576
            panelListadoCliente.Location = New Point(panelListadoCliente.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelListadoCliente.Location.X = -586
            panelListadoCliente.Location = New Point(panelListadoCliente.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub btnVehiculo_Click(sender As Object, e As EventArgs) Handles btnVehiculo.Click        'click vehiculo
        opcion = 3
        botonesOpciones()
    End Sub

    Private Sub picVehiculoGestion_Click(sender As Object, e As EventArgs) Handles picVehiculoGestion.Click        'Abrir form Vehiculo
        picVehiculoGestion.Hide()
        picVehiculoListado.Hide()
        If abierto = False Then
            abierto = True
            panelVehiculo.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelVehiculo.Location.X = 159
                panelVehiculo.Location = New Point(panelVehiculo.Location.X + 1, 0)
            Loop

            Do Until panelVehiculo.Location.X = 169
                panelVehiculo.Location = New Point(panelVehiculo.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop

        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click 'Cerrar Form Vehiculo
        abierto = False
        charmLogo.Visible = False
        panelVehiculo.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelVehiculo.Location.X = -576
            panelVehiculo.Location = New Point(panelVehiculo.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelVehiculo.Location.X = -586
            panelVehiculo.Location = New Point(panelVehiculo.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picVehiculoListado_Click(sender As Object, e As EventArgs) Handles picVehiculoListado.Click        'abrir form listado vehiculo
        picVehiculoGestion.Hide()
        picVehiculoListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoVehiculo.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoVehiculo.Location.X = 159
                panelListadoVehiculo.Location = New Point(panelListadoVehiculo.Location.X + 1, 0)
            Loop
            Do Until panelListadoVehiculo.Location.X = 169
                panelListadoVehiculo.Location = New Point(panelListadoVehiculo.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click 'cerrar form listado vehiculo
        abierto = False
        charmLogo.Visible = False
        panelListadoVehiculo.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoVehiculo.Location.X = -576
            panelListadoVehiculo.Location = New Point(panelListadoVehiculo.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelListadoVehiculo.Location.X = -586
            panelListadoVehiculo.Location = New Point(panelListadoVehiculo.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub btnCategoria_Click(sender As Object, e As EventArgs) Handles btnCategoria.Click        'click categoria
        opcion = 4
        botonesOpciones()
    End Sub

    Private Sub picCategoriaGestion_Click(sender As Object, e As EventArgs) Handles picCategoriaGestion.Click        'Abrir form categoria
        picCategoriaGestion.Hide()
        picCategoriaListado.Hide()
        If abierto = False Then
            abierto = True
            panelCategoria.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelCategoria.Location.X = 159
                panelCategoria.Location = New Point(panelCategoria.Location.X + 1, 0)
            Loop

            Do Until panelCategoria.Location.X = 169
                panelCategoria.Location = New Point(panelCategoria.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop

        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click 'cerrar form categoria 
        abierto = False
        charmLogo.Visible = False
        panelCategoria.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelCategoria.Location.X = -576
            panelCategoria.Location = New Point(panelCategoria.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop

        Do Until panelCategoria.Location.X = -586
            panelCategoria.Location = New Point(panelCategoria.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picCategoriaListado_Click(sender As Object, e As EventArgs) Handles picCategoriaListado.Click        'abrir form listado categoria
        picCategoriaGestion.Hide()
        picCategoriaListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoCategoria.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoCategoria.Location.X = 159
                panelListadoCategoria.Location = New Point(panelListadoCategoria.Location.X + 1, 0)
            Loop

            Do Until panelListadoCategoria.Location.X = 169
                panelListadoCategoria.Location = New Point(panelListadoCategoria.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click 'cerrar form listado categoria
        abierto = False
        charmLogo.Visible = False
        panelListadoCategoria.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoCategoria.Location.X = -576
            panelListadoCategoria.Location = New Point(panelListadoCategoria.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelListadoCategoria.Location.X = -586
            panelListadoCategoria.Location = New Point(panelListadoCategoria.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)  'abrir ventana config
        ventanaEstadistica.Show()
    End Sub
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click 'Añadir Categoria 
        conexion()
        Try
            cm.CommandText = "insert into categoria values (" & Chr(34) & txtNombreCategoria.Text & Chr(34) & "," & Chr(34) & txtdescripcionCategoria.Text & Chr(34) & "," &
            txtPrecioCategoria.Text & "," & Chr(34) & "Activo" & Chr(34) & ");"
            cm.ExecuteNonQuery()
            MsgBox("Categoria añadida correctamente.")
        Catch
            MsgBox("N°Error 53: Error al añadir categoria.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'actualizar la tabla en el listado
        DataGridCategoriasListado.DataSource.clear()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategoriasListado.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 54: Error al cargar categorias.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click 'Cargar Categoria
        Dim estado As String
        conexion()
        Try
            cm.CommandText = "SELECT descripcion_categoria,precio,estado FROM categoria where nombre_categoria = " & Chr(34) & txtNombreCategoria.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            txtdescripcionCategoria.Text = read.Item(0).ToString
            txtPrecioCategoria.Text = read.Item(1).ToString
            estado = read.Item(2).ToString
        Catch
            MsgBox("N°Error 55: Error al cargar categoria.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        If estado = "Baja" Then
            CheckBajaCategoria.Checked = True
        Else
            CheckBajaCategoria.Checked = False
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click 'Modificar Categoria 
        If txtNombreCategoria.Text = "" Then
            MsgBox("Ingrese valores requeridos.")
        Else
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE categoria SET descripcion_categoria= " & Chr(34) & txtdescripcionCategoria.Text & Chr(34) & 'actualiza la descripcion
                    " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 56: Error al actualizar descripcion de categoria.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE categoria SET precio= " & Chr(34) & txtPrecioCategoria.Text & Chr(34) & 'actualiza el precio
                    " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 57: Error al actualizar precio de categoria.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                If CheckBajaCategoria.Checked = True Then 'checkbox baja
                    conexion()
                    Try
                        cm.CommandText = "UPDATE categoria SET estado = 'Baja'" & 'actualiza el estado, lo da de baja
                        " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                        cm.ExecuteNonQuery()
                        MsgBox("Categoria dado de baja correctamente.")
                    Catch
                        MsgBox("N°Error 58: Error al dar de baja categoria.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                Else 'lo da de alta 
                    conexion()
                    Try
                        cm.CommandText = "UPDATE categoria SET estado = 'Activo'" & 'actualiza el estado, lo da de baja
                        " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                        cm.ExecuteNonQuery()
                    Catch
                        MsgBox("N°Error 59: Error al dar de alta categoria.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                End If
                MsgBox("Categoria actualizada correctamente.")

            Catch
                MsgBox("N°Error 60: Error al modificar categoria.")
            End Try
        End If

        'actualizar la tabla en el listado
        DataGridCategoriasListado.DataSource.clear()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategoriasListado.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 61: Error al cargar categorias.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click  'Añadir Vehiculo 
        If txtPadronVehiculo.Text = "" Or txtAñoVehiculo.Text = "" Or txtMarcaVehiculo.Text = "" Or txtModeloVehiculo.Text = "" Or
           txtMatriculavehiculo.Text = "" Or txtKilometrajeVehiculo.Text = "" Or txtColorVehiculo.Text = "" Then
            txtPadronVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtAñoVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtMarcaVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtModeloVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtMatriculavehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtKilometrajeVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtColorVehiculo.BackColor = Color.FromArgb(255, 192, 128)
        Else
            txtPadronVehiculo.BackColor = Color.White
            txtAñoVehiculo.BackColor = Color.White
            txtMarcaVehiculo.BackColor = Color.White
            txtModeloVehiculo.BackColor = Color.White
            txtMatriculavehiculo.BackColor = Color.White
            txtKilometrajeVehiculo.BackColor = Color.White
            txtColorVehiculo.BackColor = Color.White

            'año del vehiculo
            Dim año As String = txtAñoVehiculo.Text
            If año >= 1950 And año <= 2100 Then
                Dim a As Boolean = False
                conexion()
                Try
                    cm.CommandText = "insert into vehiculo values (" & Chr(34) & txtPadronVehiculo.Text & Chr(34) & "," & txtAñoVehiculo.Text & "," &
                   Chr(34) & txtMarcaVehiculo.Text & Chr(34) & "," & Chr(34) & txtModeloVehiculo.Text & Chr(34) & "," & Chr(34) & txtMatriculavehiculo.Text & Chr(34) &
                   "," & Chr(34) & txtKilometrajeVehiculo.Text & Chr(34) & "," & Chr(34) & txtCombustibleVehiculo.Text & Chr(34) & "," & Chr(34) & txtColorVehiculo.Text & Chr(34) & ", 0," &
                   Chr(34) & "Disponible" & Chr(34) & "," & Chr(34) & ComboBoxCategoriaVehiculo.Text & Chr(34) & "," & Chr(34) & "Activo" & Chr(34) & ",null" & ");"
                    cm.ExecuteNonQuery()
                    a = True
                Catch
                    MsgBox("N°Error 62: Error al añadir vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                If a = True Then
                    Dim idS As Integer
                    conexion()
                    Try
                        cm.CommandText = "SELECT id_sucursal from sucursal where nombre_sucursal = " & Chr(34) & ComboBoxSucursalV.Text & Chr(34)
                        read = cm.ExecuteReader
                        read.Read()
                        idS = read.Item(0)
                    Catch
                        MsgBox("N°Error 63: Error al obtener id de sucursal.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)

                    conexion()
                    Try
                        'inserta datos a tabla esta, que auto en que sucursal
                        cm.CommandText = "insert into esta values (" & idS & "," & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ");"
                        cm.ExecuteNonQuery()
                        MsgBox("Vehiculo añadido correctamente.")
                    Catch
                        MsgBox("N°Error 64: Error al añadir vehiculo.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                End If

                'refrescar datagrid vehiculo
                DataGridVehiculo.DataSource.clear()
                conexion()
                Try
                    cm.CommandText = "SELECT * FROM vehiculo where disponibilidad = 'Disponible'"
                    cm.ExecuteNonQuery()
                    da.SelectCommand = cm
                    da.Fill(ds, "Vehiculos")
                    DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
                Catch
                    MsgBox("N°Error 65: Error al cargar vehiculos.")
                End Try
                cx.Close()
            Else
                MsgBox("N°Error 66: Año del vehiculo no valido.")
            End If


        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click  'Cargar Vehiculo
        Dim estadoV As String
        'añadir foto del vehiculo
        Dim fotoo As String

        conexion()
        Try
            cm.CommandText = "SELECT ano,marca,modelo,matricula,combustible,kilometraje,color,nombre_categoria,estado,foto FROM vehiculo where padron = " & Chr(34) & txtPadronVehiculo.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            txtAñoVehiculo.Text = read.Item(0).ToString
            txtMarcaVehiculo.Text = read.Item(1).ToString
            txtModeloVehiculo.Text = read.Item(2).ToString
            txtMatriculavehiculo.Text = read.Item(3).ToString
            txtKilometrajeVehiculo.Text = read.Item(5).ToString
            txtColorVehiculo.Text = read.Item(6).ToString
            txtCombustibleVehiculo.SelectedIndex = txtCombustibleVehiculo.FindString(read.Item(4).ToString)
            ComboBoxCategoriaVehiculo.SelectedIndex = ComboBoxCategoriaVehiculo.FindString(read.Item(7).ToString)
            estadoV = read.Item(8).ToString
            fotoo = read.Item(9).ToString
        Catch
            MsgBox("N°Error 67: Error al cargar vehiculo.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        Try
            'añadir foto del vehiculo
            fotoo = fotoo
            Dim folder As String = Application.StartupPath & "\autos"
            Dim filename As String = System.IO.Path.Combine(folder, fotoo & ".jpg")
            fotoVehiculo.Image = Image.FromFile(filename)
        Catch
        End Try


        'dependiendo del estado del vehiculo, checkea o no el checkbox de la baja del vehiculo
        If estadoV = "Baja" Then
            CheckBoxBaja.Checked = True
        Else
            CheckBoxBaja.Checked = False
        End If

        'sucursal del vehiculo
        Dim idS As Integer
        conexion()
        Try
            cm.CommandText = "SELECT id_sucursal FROM esta where padron = " & Chr(34) & txtPadronVehiculo.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            idS = read.Item(0)
        Catch
            MsgBox("N°Error 68: Error al obtener id de sucursal.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT nombre_sucursal FROM sucursal where id_sucursal = " & idS
            read = cm.ExecuteReader
            read.Read()
            ComboBoxSucursalV.SelectedIndex = ComboBoxSucursalV.FindString(read.Item(0).ToString)
        Catch
            MsgBox("N°Error 69: Error al cargar combobox de sucursal.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click  'Modificar vehiculo
        Dim fechaBaja As String
        If CheckBoxBaja.Checked Then
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET estado = " & Chr(34) & "Baja" & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                    MsgBox(cm.CommandText)
                Catch
                    MsgBox("N°Error 70: Error al actualizar estado del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                'añadir fecha de baja
                fechaBaja = DateAndTime.Now.ToString("yyyy-MM-dd")
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET fecha_baja_vehiculo = " & Chr(34) & fechaBaja & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 71: Error al actualizar fecha de baja del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)
                MsgBox("Vehiculo dado de baja correctamente.")
            Catch
                MsgBox("N°Error 72: Error al dar de baja vehiculo.")
            End Try
        Else
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET ano = " & Chr(34) & txtAñoVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 73: Error al actualizar año del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET marca = " & Chr(34) & txtMarcaVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 74: Error al actualizar marca del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET modelo = " & Chr(34) & txtModeloVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 75: Error al actualizar modelo del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET matricula = " & Chr(34) & txtMatriculavehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 76: Error al actualizar matricula del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET nombre_categoria = " & Chr(34) & ComboBoxCategoriaVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 77: Error al actualizar categoria del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET combustible = " & Chr(34) & txtCombustibleVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 78: Error al actualizar combustible del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET kilometraje = " & Chr(34) & txtKilometrajeVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 79: Error al actualizar kilometraje del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET color = " & Chr(34) & txtColorVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 80: Error al actualizar color del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                'cambiar sucursal obtener primero el id y luego aplicar a tabla esta--
                Dim idS As Integer
                conexion()
                Try
                    cm.CommandText = "select id_sucursal from sucursal where nombre_sucursal = " & Chr(34) & ComboBoxSucursalV.Text & Chr(34)
                    read = cm.ExecuteReader
                    read.Read()
                    idS = read.Item(0)
                Catch
                    MsgBox("N°Error 81: Error al obtener id de sucursal.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE esta SET id_sucursal = " & idS &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 82: Error al actualizar sucursal del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                '--------------------------------------------------------------------

                MsgBox("Vehiculo actualizado correctamente.")
            Catch
                MsgBox("N°Error 83: Error al actualizar vehiculo.")
            End Try

        End If

        DataGridVehiculo.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM vehiculo where disponibilidad = 'Disponible'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 84: Error al cargar vehiculos.")
        End Try
        cx.Close()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click  'consulta cliente
        conexion()
        Me.DataGridListadoCliente.DataSource.clear()

        Try
            cm.CommandText = "SELECT * FROM cliente where " & ComboBoxCliente.Text & " = " & Chr(34) & txtClienteConsulta.Text & Chr(34)
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Cliente")
            DataGridListadoCliente.DataSource = ds.Tables("Cliente")
        Catch
            MsgBox("N°Error 85: Error al cargar cliente.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click  'añadir cliente

        If txtDocumentoCliente.Text = "" Or txtNombreCliente.Text = "" Or txtApellidoCliente.Text = "" Or txtTelefonoCliente.Text = "" Or
           txtPais.Text = "" Or txtCiudad.Text = "" Or txtZona.Text = "" Or txtCalle.Text = "" Or txtNumeroPuerta.Text = "" Then
            Label25.BackColor = Color.FromArgb(255, 192, 128)
            Label28.BackColor = Color.FromArgb(255, 192, 128)
            Label124.BackColor = Color.FromArgb(255, 192, 128)
            Label125.BackColor = Color.FromArgb(255, 192, 128)
            Label29.BackColor = Color.FromArgb(255, 192, 128)
            Label31.BackColor = Color.FromArgb(255, 192, 128)
            Label32.BackColor = Color.FromArgb(255, 192, 128)
            Label30.BackColor = Color.FromArgb(255, 192, 128)
            Label18.BackColor = Color.FromArgb(255, 192, 128)

        Else
            Label25.BackColor = Color.White
            Label28.BackColor = Color.White
            Label124.BackColor = Color.White
            Label125.BackColor = Color.White
            Label29.BackColor = Color.White
            Label31.BackColor = Color.White
            Label32.BackColor = Color.White
            Label30.BackColor = Color.White
            Label18.BackColor = Color.White

            Dim cedula As Boolean = False
            Dim total As Integer
            Dim numeroci As Integer
            Dim fechaBD As String
            Dim cedulac As Char() = txtDocumentoCliente.Text

            'corroborar edad
            Dim a As String
            a = DateDiff("yyyy", FechaNacCliente.Value, DateTime.Now).ToString()
            If a >= 18 Then
                'armar fecha de nacimiento 
                fechaBD = FechaNacCliente.Value.ToString("yyyy-MM-dd")

                If ComboBoxDocumento.Text = "Cedula" Then
                    Try
                        numeroci = Int32.Parse(cedulac)
                        For i = 0 To 6 Step 1
                            numeroci = Int32.Parse(cedulac(i))

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

                        MsgBox(total Mod 10 & "     " & cedulac(7))

                        If total Mod 10 + Int32.Parse(cedulac(7)) = 10 Then
                            cedula = True
                        Else
                            MsgBox("N°Error 86: Ingrese una cedula valida.")
                        End If
                    Catch
                        MsgBox("N°Error 87: Ingrese una cedula valida.")
                    End Try
                Else
                    cedula = True
                End If

                If cedula = True Then
                    conexion()
                    Try
                        cm.CommandText = "insert into cliente values (" & Chr(34) & txtDocumentoCliente.Text & Chr(34) & "," & Chr(34) & ComboBoxDocumento.Text & Chr(34) & "," &
                       Chr(34) & fechaBD & Chr(34) & "," & Chr(34) & txtNombreCliente.Text & Chr(34) & "," & Chr(34) & txtApellidoCliente.Text & Chr(34) &
                       "," & Chr(34) & txtTelefonoCliente.Text & Chr(34) & "," & Chr(34) & "Inactivo" & Chr(34) & "," & Chr(34) & txtZona.Text & Chr(34) &
                       "," & Chr(34) & txtCalle.Text & Chr(34) & "," & Chr(34) & txtCiudad.Text & Chr(34) &
                       "," & txtNumeroPuerta.Text & "," & Chr(34) & txtPais.Text & Chr(34) & ");"
                        cm.ExecuteNonQuery()
                        MsgBox("Cliente añadido con exito.")
                    Catch
                        MsgBox("N°Error 88: Error al añadir cliente.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)

                End If

                'refrescar tabla cliente
                DataGridListadoCliente.DataSource.clear()
                conexion()
                Try
                    cm.CommandText = "SELECT * FROM cliente"
                    cm.ExecuteNonQuery()
                    da.SelectCommand = cm
                    da.Fill(ds, "Cliente")
                    DataGridListadoCliente.DataSource = ds.Tables("Cliente")
                Catch
                    MsgBox("N°Error 89: Error al cargar clientes.")
                End Try
                cx.Close()
            Else
                MsgBox("N°Error 90: Cliente debe ser mayor de edad.")
            End If

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click 'cargar cliente
        conexion()
        Try
            cm.CommandText = "SELECT nombre_conductor,apellido_conductor,nombre_ciudad,nombre_calle,nombre_zona," &
            "numeropuerta,fecha_nacimiento,telefono,pais,estado FROM cliente where documento_cliente = " & Chr(34) & txtDocumentoCliente.Text & Chr(34) &
            "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            txtNombreCliente.Text = read.Item(0).ToString
            txtApellidoCliente.Text = read.Item(1).ToString
            txtCiudad.Text = read.Item(2).ToString
            txtCalle.Text = read.Item(3).ToString
            txtZona.Text = read.Item(4).ToString
            txtNumeroPuerta.Text = read.Item(5).ToString
            FechaNacCliente.Value = Convert.ToDateTime(read.Item(6)) 'fecha Nacimiento
            txtTelefonoCliente.Text = read.Item(7).ToString
            txtPais.Text = read.Item(8).ToString

            If read.Item(9).ToString = "Baja" Then
                CheckBoxCliente.Checked = True
            Else
                CheckBoxCliente.Checked = False
            End If
        Catch
            MsgBox("N°Error 91: Error al cargar cliente.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click  'modificar cliente
        Dim estado As String

        conexion()
        Try
            cm.CommandText = "SELECT estado from cliente where documento_cliente =" & Chr(34) & txtDocumentoCliente.Text & Chr(34) &
                    "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            estado = read.Item(0).ToString
        Catch
            MsgBox("N°Error 92: Error al obtener estado del cliente.")
        End Try
        cx.Close()
        historial(cm.CommandText)


        If estado = "Baja" And CheckBoxCliente.Checked = False Then
            conexion()
            Try
                cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Inactivo" & Chr(34) &
            " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34) & ";"
                cm.ExecuteNonQuery()
                MsgBox("Cliente dado de alta correctamente.")
            Catch
                MsgBox("N°Error 93: Error al dar de alta cliente.")
            End Try
            cx.Close()
            historial(cm.CommandText)

        End If

        If CheckBoxCliente.Checked Then 'dar de baja al cliente
            If estado = "Activo" Then
                MsgBox("N°Error 94: No se puede dar de baja a un cliente que esté alquilando.")
            Else
                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Baja" & Chr(34) &
                " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                    MsgBox("Cliente dado de baja correctamente.")
                Catch
                    MsgBox("N°Error 95: Error al dar de baja cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

            End If
        Else
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_conductor = " & Chr(34) & txtNombreCliente.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 96: Error al actualizar nombre del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET apellido_conductor = " & Chr(34) & txtApellidoCliente.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 97: Error al actualizar apellido del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_ciudad = " & Chr(34) & txtCiudad.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 98: Error al actualizar ciudad del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_calle = " & Chr(34) & txtCalle.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 99: Error al actualizar calle del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_zona = " & Chr(34) & txtZona.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 100: Error al actualizar zona del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET numeropuerta = " & Chr(34) & txtNumeroPuerta.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 101: Error al actualizar numero de puerta del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET fecha_nacimiento = " & Chr(34) & FechaNacCliente.Value.ToString("yyyy-MM-dd") & Chr(34) & 'fecha
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 102: Error al actualizar fecha de nacimiento del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET telefono = " & Chr(34) & txtTelefonoCliente.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 103: Error al actualizar telefono del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET pais = " & Chr(34) & txtPais.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 104: Error al actualizar pais del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                MsgBox("Cliente actualizado correctamente.")
            Catch
                MsgBox("N°Error 105: Error al actualizar cliente.")
            End Try

        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click 'calcular precio
        Dim precio As String = 0
        Dim diferencia As Long
        Dim d1 As Date = DateTimePickerFIN.Value
        Dim d2 As Date = DateTimePickerINI.Value

        diferencia = DateDiff("d", d2, d1) 'diferencia de dias entre 2 fechas  

        If diferencia <= 0 Then
            Label20.Text = "N°Error 106: Fechas no validas."
        Else
            conexion()
            Try
                cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & ComboBoxCategoria.Text & Chr(34)
                cm.ExecuteNonQuery()
                read = cm.ExecuteReader
                read.Read()
                precio = read.Item(0).ToString
                Label20.Text = precio * diferencia
                precio = 0
            Catch
                MsgBox("N°Error 107: Error al cargar precio de la categoria.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'Reservar
        Dim id, documento, tipoDocumento, padron, FInicio, FFinal, documentoSecundario, tipoDocSecundario, sucursal As String
        Dim diferencia As Integer
        Dim erro As Boolean = False

        documento = txtDocumento.Text
        tipoDocumento = ComboBoxTipoDeDocumento.Text
        padron = txtPadron.Text
        documentoSecundario = txtDoc2do.Text
        tipoDocSecundario = ComboBoxTipoDoc2do.Text

        'nombre de sucursal
        sucursal = ComboBoxSucursal.Text

        'Genera las fechas
        FInicio = DateTimeFechaInicioReserva.Value.ToString("yyyy-MM-dd")
        FFinal = DateTimeFechaFinReserva.Value.ToString("yyyy-MM-dd")

        'diferencia de dias de fechas
        diferencia = DateDiff("d", DateTimeFechaInicioReserva.Value, DateTimeFechaFinReserva.Value)

        'control de errores 
        If ComboBoxTipoDeDocumento.Text = "Cedula" Then 'controla la cedula del conductor
            '--------------algoritmo comprobacion de CI-------------
            Dim cedula As Char() = txtDocumento.Text
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
                    'cedula correcta
                    erro = False
                Else
                    erro = True
                    MsgBox("N°Error 108: Ingrese cedula valida.")
                End If
            Catch
                erro = True
                MsgBox("N°Error 109: Ingrese cedula valida.")
            End Try
        End If

        If ComboBoxTipoDoc2do.Text = "Cedula" Then 'controla la cedula del segundo conductor 
            '--------------algoritmo comprobacion de CI-------------
            Dim cedula As Char() = txtDoc2do.Text
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
                    'cedula correcta
                    erro = False
                Else
                    erro = True
                    MsgBox("N°Error 110: Ingrese cedula valida del segundo conductor.")
                End If
            Catch
                erro = True
                MsgBox("N°Error 111: Ingrese cedula valida del segundo conductor.")
            End Try
        End If

        'comprobar estado del vehiculo
        Dim estadoVehiculo As String
        conexion()
        Try
            cm.CommandText = "SELECT disponibilidad FROM vehiculo where padron = " & Chr(34) & txtPadron.Text & Chr(34)
            cm.ExecuteNonQuery()
            read = cm.ExecuteReader
            read.Read()
            estadoVehiculo = read.Item(0).ToString
        Catch
            MsgBox("N°Error 112: Error al cargar estado del vehiculo.")
        End Try
        cx.Close()

        If Not estadoVehiculo = "Disponible" Then
            erro = True
            MsgBox("N°Error 113: Vehiculo no disponible.")
        End If

        'comprobar fechas 
        If DateTimeFechaInicioReserva.Value < DateTime.Today Then
            erro = True
            MsgBox("N°Error 114: Ingrese fechas correctas.")
        Else
            If DateTimeFechaInicioReserva.Value > DateTimeFechaFinReserva.Value Then
                erro = True
                MsgBox("N°Error 115: Ingrese fechas correctas.")
            End If
        End If

        'comprobar campos vacios

        If ComboBoxTipoDeDocumento.Text = "" Or txtDocumento.Text = "" Or txtPadron.Text = "" Then
            erro = True
            Label7.BackColor = Color.FromArgb(255, 192, 128)
            Label6.BackColor = Color.FromArgb(255, 192, 128)
            Label111.BackColor = Color.FromArgb(255, 192, 128)
            MsgBox("N°Error 116: Llene los campos obligatorios.")
        End If

        If CheckBoxConsuctorSec.Checked = True Then
            If ComboBoxTipoDoc2do.Text = "" Or txtDoc2do.Text = "" Then
                erro = True
                LabelDoc2.BackColor = Color.FromArgb(255, 192, 128)
                labelDoc2a.BackColor = Color.FromArgb(255, 192, 128)
                MsgBox("N°Error 117: Llene los campos obligatorios.")
            End If
        End If

        If erro = False Then
            'devuelve color a normal
            Label7.BackColor = Color.White
            Label6.BackColor = Color.White
            Label111.BackColor = Color.White
            LabelDoc2.BackColor = Color.White
            labelDoc2a.BackColor = Color.White

            BoletaReserva.generarBoleta(documento, tipoDocumento, padron, FInicio, FFinal, documentoSecundario, tipoDocSecundario, diferencia, sucursal) 'generar boleta
            BoletaReserva.Show()
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click 'Listado Reserva
        Me.DataGridListadoReserva.DataSource.clear()

        If ComboBoxListadoReserva.Text = "Cliente" Then
            conexion()
            Try
                cm.CommandText = "SELECT * FROM reserva where documento_cliente = " & Chr(34) & txtDocumentoListadoRes.Text & Chr(34) & " and tipo_documento =" &
                    Chr(34) & ComboBoxDocumentoLReserva.Text & Chr(34) & ";"
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Reserva")
                DataGridListadoReserva.DataSource = ds.Tables("Reserva")
            Catch
                MsgBox("N°Error 118: Error al visualizar reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

        ElseIf ComboBoxListadoReserva.Text = "Padrón" Then
            conexion()
            Try
                cm.CommandText = "SELECT * FROM reserva where padron = " & Chr(34) & txtBuscadorListadoReserva.Text & Chr(34)
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Reserva")
                DataGridListadoReserva.DataSource = ds.Tables("Reserva")
            Catch
                MsgBox("N°Error 119: Error al visualizar reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

        ElseIf ComboBoxListadoReserva.Text = "Estado de reserva" Then
            conexion()
            Try
                cm.CommandText = "SELECT * FROM reserva where estado_reserva = " & Chr(34) & txtBuscadorListadoReserva.Text & Chr(34)
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Reserva")
                DataGridListadoReserva.DataSource = ds.Tables("Reserva")
            Catch
                MsgBox("N°Error 120: Error al visualizar reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

        ElseIf ComboBoxListadoReserva.Text = "Fecha" Then
            Dim fechaIni, fechaFin As String

            fechaIni = DateTimeIniListadoReserva.Value.ToString("yyyy-MM-dd")
            fechaFin = DateTimeFinListadoReserva.Value.ToString("yyyy-MM-dd")

            conexion()
            Try 'con fechas
                cm.CommandText = "SELECT cliente.documento_cliente,cliente.nombre_conductor,cliente.apellido_conductor,vehiculo.padron,vehiculo.matricula,vehiculo.marca,vehiculo.modelo FROM reserva,cliente,esta,vehiculo where reserva.fecha_inicio >= " & Chr(34) & fechaIni & Chr(34) & " and reserva.fecha_final <= " & Chr(34) & fechaFin & Chr(34) &
                    " and cliente.documento_cliente = reserva.documento_cliente and esta.padron = vehiculo.padron"
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Reserva")
                DataGridListadoReserva.DataSource = ds.Tables("Reserva")
            Catch
                MsgBox("N°Error 121: Error al visualizar reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------------

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxListadoReserva.SelectedIndexChanged 'al cambiar opcion dentro del combobox Listado Reserva
        lblDoc.Visible = False
        lblTipoDoc.Visible = False
        txtDocumentoListadoRes.Visible = False
        ComboBoxDocumentoLReserva.Visible = False
        lblFechaINI.Visible = False
        lblFechaFIN.Visible = False
        DateTimeIniListadoReserva.Visible = False
        DateTimeFinListadoReserva.Visible = False

        If ComboBoxListadoReserva.Text = "Cliente" Then
            lblDoc.Visible = True
            lblTipoDoc.Visible = True
            txtDocumentoListadoRes.Visible = True
            ComboBoxDocumentoLReserva.Visible = True
        ElseIf ComboBoxListadoReserva.Text = "Fecha" Then
            lblFechaINI.Visible = True
            lblFechaFIN.Visible = True
            DateTimeIniListadoReserva.Visible = True
            DateTimeFinListadoReserva.Visible = True
        End If
    End Sub

    Private Sub CheckBoxConsuctorSec_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxConsuctorSec.CheckedChanged
        If CheckBoxConsuctorSec.Checked = True Then
            ComboBoxTipoDoc2do.Visible = True
            LabelDoc2.Visible = True
            labelDoc2a.Visible = True
            txtDoc2do.Visible = True
        Else
            ComboBoxTipoDoc2do.Visible = False
            LabelDoc2.Visible = False
            labelDoc2a.Visible = False
            txtDoc2do.Visible = False
        End If
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click  'refrescar consulta de vehiculo

        If DataGridVehiculo.RowCount > 0 Then
            DataGridVehiculo.DataSource.clear()
        End If

        conexion()
        Try
            cm.CommandText = "SELECT * FROM vehiculo where disponibilidad = 'Disponible'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 122: Error al cargar vehiculos.")
        End Try
        cx.Close()

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click  'recarga listado del cliente
        DataGridListadoCliente.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM cliente"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Cliente")
            DataGridListadoCliente.DataSource = ds.Tables("Cliente")
        Catch
            MsgBox("N°Error 123: Error al cargar clientes.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click 'recargar listado reserva
        DataGridListadoReserva.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM reserva"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "rese")
            DataGridListadoReserva.DataSource = ds.Tables("rese")
        Catch
            MsgBox("N°Error 124: Error al cargar reservas.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click 'modificar reserva
        'depende del estaod de la reserva si se puede modificar o no
        Dim estado, padronAnterior As String

        conexion()
        Try
            cm.CommandText = "SELECT estado_reserva FROM reserva where id_reserva = " & txtIdReserva.Text
            read = cm.ExecuteReader
            read.Read()
            estado = read.Item(0).ToString
        Catch
        End Try

        If estado = "Reservado" Then
            'aca el codigo de modificacion

            Try
                '--------------------------cambiar estado del vehiculo ya reservado-------------------
                conexion()
                Try
                    cm.CommandText = "select padron from reserva where id_reserva = " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                    read.Read()
                    padronAnterior = read.Item(0).ToString
                Catch
                    MsgBox("N°Error 125: Error al obtener padron del vehiculo reservado.")
                End Try
                cx.Close()

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Disponible" & Chr(34) &
                    " WHERE padron= " & Chr(34) & padronAnterior & Chr(34) 'cambiar estado del anterior vehiculo a disponible 
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 126: Error al actualizar disponibilidad del vehiculo reservado.")
                End Try
                cx.Close()


                '------------------------------------------------------------------------------------
                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET padron = " & Chr(34) & txtPadron.Text & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 127: Error al actualizar padron de la reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Reservado" & Chr(34) &
                    " WHERE padron= " & txtPadron.Text 'cambiar estado del nuevo a Reservado
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 128: Error al actualizar disponibilidad del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET fecha_inicio = " & Chr(34) & DateTimeFechaInicioReserva.Value.ToString("yyyy-MM-dd") & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 129: Error al actualizar fecha inicial de reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET fecha_final = " & Chr(34) & DateTimeFechaFinReserva.Value.ToString("yyyy-MM-dd") & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 130: Error al actualizar fecha final de reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                'actualizar la tabla está
                Dim idS As Integer
                conexion()
                Try
                    cm.CommandText = "select id_sucursal from sucursal where nombre_sucursal = " & Chr(34) & ComboBoxSucursal.Text & Chr(34)
                    cm.ExecuteNonQuery()
                    read.Read()
                    idS = read.Item(0)
                Catch
                    MsgBox("N°Error 131: Error al cargar id de sucursal.")
                End Try
                cx.Close()

                conexion()
                Try
                    cm.CommandText = "UPDATE esta SET id_sucursal = " & idS &
                    " WHERE padron= " & txtPadron.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 132: Error al actualizar id de sucursal del vehiculo.")
                End Try
                cx.Close()

                '--------------------------------------
                If CheckBoxConsuctorSec.Checked = True Then
                    conexion()
                    Try
                        cm.CommandText = "UPDATE reserva SET documento_2doconductor = " & Chr(34) & txtDoc2do.Text & Chr(34) &
                        " WHERE id_reserva= " & txtIdReserva.Text
                        cm.ExecuteNonQuery()
                    Catch
                        MsgBox("N°Error 133: Error al actualizar documento del segundo conductor en la reserva.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)

                    conexion()
                    Try
                        cm.CommandText = "UPDATE reserva SET tipo_documento2do = " & Chr(34) & ComboBoxTipoDoc2do.Text & Chr(34) &
                        " WHERE id_reserva= " & txtIdReserva.Text
                        cm.ExecuteNonQuery()
                    Catch
                        MsgBox("N°Error 134: Error al actualizar el tipo de documento del segundo conductor en la reserva.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)

                End If

                MsgBox("Reserva actualizada correctamente.")
            Catch
                MsgBox("N°Error 135: Error al actualizar reserva.")
            End Try

        Else
            MsgBox("N°Error 136: No se puede modificar un alquiler.")
        End If

    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click 'cargar reserva
        Dim estado As String
        Dim erro As Boolean = True
        conexion()
        Try
            cm.CommandText = "SELECT documento_cliente,tipo_documento,padron,fecha_inicio,fecha_final,fecha_retiro,fecha_devolucion" &
                ", estado_reserva, documento_2doconductor, tipo_documento2do FROM reserva where id_reserva = " & txtIdReserva.Text
            read = cm.ExecuteReader
            read.Read()
            txtDocumento.Text = read.Item(0).ToString
            ComboBoxTipoDeDocumento.SelectedIndex = ComboBoxTipoDeDocumento.FindString(read.Item(1).ToString)
            txtPadron.Text = read.Item(2).ToString
            DateTimeFechaInicioReserva.Value = Convert.ToDateTime(read.Item(3))
            DateTimeFechaFinReserva.Value = Convert.ToDateTime(read.Item(4))

            If Not IsDBNull(read.Item(5)) Then
                DateTimePicker4.Value = Convert.ToDateTime(read.Item(5))
            End If

            If Not IsDBNull(read.Item(6)) Then
                DateTimePicker3.Value = Convert.ToDateTime(read.Item(6))
            End If

            estado = read.Item(7).ToString

            If Not "" = read.Item(8).ToString Then
                CheckBoxConsuctorSec.Checked = True
                txtDoc2do.Text = read.Item(8).ToString
                ComboBoxTipoDoc2do.SelectedIndex = ComboBoxTipoDoc2do.FindString(read.Item(9).ToString)
            Else
                CheckBoxConsuctorSec.Checked = False
            End If
            erro = False
        Catch
            MsgBox("N°Error 137: Error al cargar reserva.")
        End Try
        cx.Close()
        historial(cm.CommandText)


        If erro = False Then

            'cargar la sucursal del vehiculo
            Dim ids As Integer
            Dim erro2 As Boolean = True
            conexion()
            Try
                cm.CommandText = "SELECT id_sucursal from esta where padron = " & Chr(34) & txtPadron.Text & Chr(34)
                read = cm.ExecuteReader
                read.Read()
                ids = read.Item(0)
                erro2 = False
            Catch
                MsgBox("N°Error 138: Error al obtener id del sucursal.")
            End Try
            cx.Close()
            If erro2 = False Then
                conexion()
                Try
                    cm.CommandText = "SELECT nombre_sucursal from sucursal where id_sucursal = " & ids
                    read = cm.ExecuteReader
                    read.Read()
                    ComboBoxSucursal.SelectedIndex = ComboBoxSucursal.FindString(read.Item(0).ToString)
                    ComboBoxSucursal.Show()
                    Label12.Show()
                Catch
                    MsgBox("N°Error 139: Error al obtener nombre del sucursal.")
                End Try
                cx.Close()
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 'alquilar
        Dim idRes, FRetiro, sucursal As String
        Dim diferencia As Integer
        'fecha retiro
        FRetiro = DateTimePicker4.Value.ToString("yyyy-MM-dd")

        'nombre de sucursal
        sucursal = ComboBoxSucursal.Text

        'diferencia de dias de fechas
        diferencia = DateDiff("d", DateTimeFechaInicioReserva.Value, DateTimeFechaFinReserva.Value)

        idRes = txtIdReserva.Text
        BoletaReserva.generarBoletaAlquiler(idRes, diferencia, FRetiro, sucursal) 'generar boleta
        BoletaReserva.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click 'finalizar alquiler
        Dim finalizado As Boolean = True
        Dim idRes, FFinal, sucursal, alqui As String
        Dim diferencia As Integer
        Dim erro As Boolean = True

        'obtener estado de alquiler para validar
        conexion()
        Try
            cm.CommandText = "SELECT estado_reserva FROM reserva where id_reserva = " & txtIdReserva.Text & ";"
            read = cm.ExecuteReader
            read.Read()
            alqui = read.Item(0).ToString
            erro = False
        Catch
            MsgBox("Alquiler ya finalizado.")
            erro = True
        End Try
        cx.Close()

        'fecha retiro
        FFinal = DateTimePicker3.Value.ToString("yyyy-MM-dd")

        'nombre de sucursal
        sucursal = ComboBoxSucursal.Text

        'fijar si la fecha de devolucion esta despues de la de retiro
        If DateTimePicker3.Value < DateTimePicker4.Value Then
            MsgBox("N°Error 293: Error Con fecha de devolucion.")
            erro = True
        Else
            erro = False
        End If

        If alqui = "Finalizado" Then
            MsgBox("Alquiler ya finalizado.")
            erro = True
        End If

        'diferencia de dias de fechas
        diferencia = DateDiff("d", DateTimeFechaInicioReserva.Value, DateTimeFechaFinReserva.Value)
        If erro = False Then
            idRes = txtIdReserva.Text
            'abrir boleta fin alquiler  
            BoletaReserva.generarBoletaFinAlquiler(idRes, diferencia, FFinal, sucursal)
            BoletaReserva.Show()
        End If
    End Sub

    Private Sub btnCheck_Click(sender As Object, e As EventArgs) Handles btnCheck.Click 'check vehiculo en la recepcion
        'verifica disponibilidad del vehiculo
        Dim disp, estado As String
        conexion()
        Try
            cm.CommandText = "SELECT disponibilidad, estado from vehiculo where padron = " & Chr(34) & txtPadron.Text & Chr(34) 'comprobar esto!!
            read = cm.ExecuteReader
            read.Read()
            disp = read.Item(0).ToString
            estado = read.Item(1).ToString
        Catch
            MsgBox("N°Error 140: Error al obtener estado y disponibilidad del vehiculo.")
        End Try
        cx.Close()

        If estado = "Activo" And disp = "Disponible" Then
            'obtener nombre sucursal del vehiculo
            'cargar el combobox con el sucursal
            Dim idS As Integer
            conexion()
            Try
                cm.CommandText = "SELECT id_sucursal from esta where padron = " & Chr(34) & txtPadron.Text & Chr(34)
                read = cm.ExecuteReader
                read.Read()
                idS = read.Item(0)
            Catch
                MsgBox("N°Error 141: Error al obtener id de la sucursal.")
            End Try
            cx.Close()

            conexion()
            Try
                cm.CommandText = "SELECT nombre_sucursal from sucursal where id_sucursal = " & idS
                read = cm.ExecuteReader
                read.Read()
                ComboBoxSucursal.SelectedIndex = ComboBoxSucursal.FindString(read.Item(0).ToString)
            Catch
                MsgBox("N°Error 142: Error al obtener nombre de la sucursal.")
            End Try
            cx.Close()

            'mostrar los elementos de la interfaz
            DateTimeFechaInicioReserva.Show()
            DateTimeFechaFinReserva.Show()
            ComboBoxSucursal.Show()
            CheckBoxConsuctorSec.Show()
            Label12.Show()
            Label9.Show()
            Label10.Show()

        Else
            MsgBox("N°Error 143: Vehiculo no disponible.")
        End If
    End Sub

    Private Sub txtAñoVehiculo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAñoVehiculo.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtKilometrajeVehiculo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtKilometrajeVehiculo.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPrecioCategoria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrecioCategoria.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtTelefonoCliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTelefonoCliente.KeyPress, txtNumeroPuerta.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumeroPuerta_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button21_Click_1(sender As Object, e As EventArgs) Handles ImagenVehiculo.Click  'consulta vehiculo
        Me.DataGridVehiculo.DataSource.clear()
        If ComboBoxVehiculo.Text = "sucursal" Then 'busca por sucursal
            conexion()
            Try
                cm.CommandText = "SELECT * FROM vehiculo,esta where esta.id_sucursal = " & BuscadorVehiculo.Text & "and esta.padron = vehiculo.padron;"
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Vehiculos")
                DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
            Catch
                MsgBox("N°Error 144: Error al cargar Vehiculos de la sucursal.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        Else 'busca por los demas
            conexion()
            Try
                cm.CommandText = "SELECT * FROM vehiculo where " & ComboBoxVehiculo.Text & " = " & Chr(34) & BuscadorVehiculo.Text & Chr(34)
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Vehiculos")
                DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
            Catch
                MsgBox("N°Error 145: Error al cargar Vehiculos.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub txtNombreCliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNombreCliente.KeyPress, txtZona.KeyPress, txtPais.KeyPress, txtCiudad.KeyPress, txtCalle.KeyPress, txtApellidoCliente.KeyPress, txtColorVehiculo.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub
    '-----------------------------------------------------------------------------------------------------------------------------------------------

End Class