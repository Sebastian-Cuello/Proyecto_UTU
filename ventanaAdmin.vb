Imports System.Data.Odbc
Public Class ventanaAdmin
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
    Dim caso5 As Boolean = False
    Dim caso6 As Boolean = False
    Dim caso7 As Boolean = False
    Dim opcion As Integer 'opcion del menu para desplegar submenus

    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=" & LoginSU.user & ";PWD=" & LoginSU.contraseña
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 146: No hay conexion con la base de datos.")
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
            MsgBox("N°Error 147: Error al mantener historial.")
        End Try
        cx.Close()
    End Sub

    Private Sub pictureCerrarSesion_Click(sender As Object, e As EventArgs) Handles pictureCerrarSesion.Click  'flecha cerrar sesion
        Me.Close()
        LoginSU.Show()
    End Sub

    Private Sub ventanaAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            MsgBox("N°Error 148: Error al cargar sucursales.")
        End Try
        cx.Close()

        conexion() 'carga datagrid cliente
        Try
            cm.CommandText = "SELECT * FROM cliente"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Cliente")
            DataGridListadoCliente.DataSource = ds.Tables("Cliente")
        Catch
            MsgBox("N°Error 149: Error al cargar clientes.")
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
            MsgBox("N°Error 150: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM sucursal" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Sucursal")
            DataGridSucursal.DataSource = ds.Tables("Sucursal")
            DataGridView8.DataSource = ds.Tables("Sucursal")
        Catch
            MsgBox("N°Error 151: Error al cargar sucursales.")
        End Try
        cx.Close()

        conexion()
        Try 'carga todas las caracteristicas a su grid
            cm.CommandText = "SELECT * FROM caracteristica"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas")
            DataGridTodasCaracteristicas.DataSource = ds.Tables("Caracteristicas")
        Catch
            MsgBox("N°Error 152: Error al cargar caracteristicas.")
        End Try
        cx.Close()

        conexion()
        Try 'carga todas los mantenimientos a su grid
            cm.CommandText = "SELECT * FROM mantenimiento"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Mantenimientos")
            DataGridTodosLosMantenimientos.DataSource = ds.Tables("Mantenimientos")
        Catch
            MsgBox("N°Error 153: Error al cargar mantenimientos.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategorias.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 154: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion() 'carga vehiculos a su grid
        Try
            cm.CommandText = "SELECT * FROM vehiculo"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 155: Error al cargar vehiculos.")
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
            MsgBox("N°Error 156: Error al cargar reservas.")
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
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasGestion.Hide()
                    picMantenimientoGestion.Hide()
                    picMantenimientoGestion.Hide()
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
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasGestion.Hide()
                    picMantenimientoGestion.Hide()
                    picMantenimientoGestion.Hide()
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
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasGestion.Hide()
                    picMantenimientoGestion.Hide()
                    picMantenimientoGestion.Hide()
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
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
                    picMantenimientoGestion.Hide()
                    picMantenimientoListado.Hide()
                    caso4 = True
                Else
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso4 = False
                End If
            Case 5
                If caso5 = False Then
                    picMantenimientoGestion.Show()
                    picMantenimientoListado.Show()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    caso5 = True
                Else
                    picMantenimientoGestion.Hide()
                    picMantenimientoListado.Hide()
                    caso5 = False
                End If
            Case 6
                If caso6 = False Then
                    picCaracteristicasGestion.Show()
                    picCaracteristicasListado.Show()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    picMantenimientoGestion.Hide()
                    picMantenimientoListado.Hide()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    caso6 = True
                Else
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
                    caso6 = False
                End If
            Case 7
                If caso7 = False Then
                    picSucursalesGestion.Show()
                    picSucursalesListado.Show()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
                    picCategoriaGestion.Hide()
                    picCategoriaListado.Hide()
                    picMantenimientoListado.Hide()
                    picMantenimientoGestion.Hide()
                    picReservaGestion.Hide()
                    picReservaListado.Hide()
                    picClienteListado.Hide()
                    picClienteGestion.Hide()
                    picVehiculoGestion.Hide()
                    picVehiculoListado.Hide()
                    caso7 = True
                Else
                    picSucursalesGestion.Hide()
                    picSucursalesListado.Hide()
                    caso7 = False
                End If
        End Select
    End Sub

    Private Sub btnReserva_Click(sender As Object, e As EventArgs) Handles btnReserva.Click  'click boton reserva
        opcion = 1
        botonesOpciones()
    End Sub

    Private Sub picReservaGestion_Click(sender As Object, e As EventArgs) Handles picReservaGestion.Click  'Abrir form gestion reserva
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

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click  'Cerrar Formulario gestion Reserva
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

    Private Sub picReservaListado_Click(sender As Object, e As EventArgs) Handles picReservaListado.Click  'Abrir listado Reserva
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

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click  'Cerrar listado Reserva
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

    Private Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click  'click boton cliente
        opcion = 2
        botonesOpciones()
    End Sub

    Private Sub picClienteGestion_Click(sender As Object, e As EventArgs) Handles picClienteGestion.Click  'Abrir form cliente
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
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click  'Cerrar form cliente
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

    Private Sub picClienteListado_Click(sender As Object, e As EventArgs) Handles picClienteListado.Click  'abrir Form listado cliente
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

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click  'cerrar form listado cliente
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

    Private Sub btnVehiculo_Click(sender As Object, e As EventArgs) Handles btnVehiculo.Click  'click vehiculo
        opcion = 3
        botonesOpciones()
    End Sub

    Private Sub picVehiculoGestion_Click(sender As Object, e As EventArgs) Handles picVehiculoGestion.Click  'Abrir form Vehiculo
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

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click  'Cerrar Form Vehiculo
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

    Private Sub picVehiculoListado_Click(sender As Object, e As EventArgs) Handles picVehiculoListado.Click  'abrir form listado vehiculo
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

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click  'cerrar form listado vehiculo
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

    Private Sub btnCategoria_Click(sender As Object, e As EventArgs) Handles btnCategoria.Click  'click categoria
        opcion = 4
        botonesOpciones()
    End Sub

    Private Sub picCategoriaGestion_Click(sender As Object, e As EventArgs) Handles picCategoriaGestion.Click  'Abrir form categoria
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

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click  'cerrar form categoria 
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

    Private Sub picCategoriaListado_Click(sender As Object, e As EventArgs) Handles picCategoriaListado.Click  'abrir form listado categoria
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

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click  'cerrar form listado categoria
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

    Private Sub btnMantenimiento_Click(sender As Object, e As EventArgs) Handles btnMantenimiento.Click  'click mantenimiento
        opcion = 5
        botonesOpciones()
    End Sub

    Private Sub picMantenimientoGestion_Click(sender As Object, e As EventArgs) Handles picMantenimientoGestion.Click  'abrir form mantenimiento
        picMantenimientoGestion.Hide()
        picMantenimientoListado.Hide()
        If abierto = False Then
            abierto = True
            panelMantenimiento.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelMantenimiento.Location.X = 159
                panelMantenimiento.Location = New Point(panelMantenimiento.Location.X + 1, 0)
            Loop

            Do Until panelMantenimiento.Location.X = 169
                panelMantenimiento.Location = New Point(panelMantenimiento.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click  'cerrar form manenimiento
        abierto = False
        charmLogo.Visible = False
        panelMantenimiento.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelMantenimiento.Location.X = -576
            panelMantenimiento.Location = New Point(panelMantenimiento.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelMantenimiento.Location.X = -586
            panelMantenimiento.Location = New Point(panelMantenimiento.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picMantenimientoListado_Click(sender As Object, e As EventArgs) Handles picMantenimientoListado.Click  'Abrir form Listado Mantenimiento
        picMantenimientoGestion.Hide()
        picMantenimientoListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoMantenimiento.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoMantenimiento.Location.X = 159
                panelListadoMantenimiento.Location = New Point(panelListadoMantenimiento.Location.X + 1, 0)
            Loop

            Do Until panelListadoMantenimiento.Location.X = 169
                panelListadoMantenimiento.Location = New Point(panelListadoMantenimiento.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click  'cerrar form listado mantenimiento
        abierto = False
        charmLogo.Visible = False
        panelListadoMantenimiento.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoMantenimiento.Location.X = -576
            panelListadoMantenimiento.Location = New Point(panelListadoMantenimiento.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelListadoMantenimiento.Location.X = -586
            panelListadoMantenimiento.Location = New Point(panelListadoMantenimiento.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub btnCaracteristica_Click(sender As Object, e As EventArgs) Handles btnCaracteristica.Click  'click caracteristicas
        opcion = 6
        botonesOpciones()
    End Sub

    Private Sub picCaracteristicasGestion_Click(sender As Object, e As EventArgs) Handles picCaracteristicasGestion.Click  'abrir form caracteristica
        picCaracteristicasGestion.Hide()
        picCaracteristicasListado.Hide()
        If abierto = False Then
            abierto = True
            panelCaracteristicas.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelCaracteristicas.Location.X = 159
                panelCaracteristicas.Location = New Point(panelCaracteristicas.Location.X + 1, 0)
            Loop

            Do Until panelCaracteristicas.Location.X = 169
                panelCaracteristicas.Location = New Point(panelCaracteristicas.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles PictureBox18.Click  'Cerrar form caracteristicas
        abierto = False
        charmLogo.Visible = False
        panelCaracteristicas.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelCaracteristicas.Location.X = -576
            panelCaracteristicas.Location = New Point(panelCaracteristicas.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelCaracteristicas.Location.X = -586
            panelCaracteristicas.Location = New Point(panelCaracteristicas.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picCaracteristicasListado_Click(sender As Object, e As EventArgs) Handles picCaracteristicasListado.Click  'abrir form listado caracteristicas
        picCaracteristicasGestion.Hide()
        picCaracteristicasListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoCaracteristica.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoCaracteristica.Location.X = 159
                panelListadoCaracteristica.Location = New Point(panelListadoCaracteristica.Location.X + 1, 0)
            Loop

            Do Until panelListadoCaracteristica.Location.X = 169
                panelListadoCaracteristica.Location = New Point(panelListadoCaracteristica.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles PictureBox19.Click  'cerrar form listado caracteristicas
        abierto = False
        charmLogo.Visible = False
        panelListadoCaracteristica.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoCaracteristica.Location.X = -576
            panelListadoCaracteristica.Location = New Point(panelListadoCaracteristica.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelListadoCaracteristica.Location.X = -586
            panelListadoCaracteristica.Location = New Point(panelListadoCaracteristica.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub btnSucursales_Click(sender As Object, e As EventArgs) Handles btnSucursales.Click  'click sucursales
        opcion = 7
        botonesOpciones()
    End Sub

    Private Sub picSucursalesGestion_Click(sender As Object, e As EventArgs) Handles picSucursalesGestion.Click  'abrir form sucursales
        picSucursalesGestion.Hide()
        picSucursalesListado.Hide()
        If abierto = False Then
            abierto = True
            panelSucursal.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelSucursal.Location.X = 159
                panelSucursal.Location = New Point(panelSucursal.Location.X + 1, 0)
            Loop

            Do Until panelSucursal.Location.X = 169
                panelSucursal.Location = New Point(panelSucursal.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles PictureBox20.Click  'Cerrar form sucursal
        abierto = False
        charmLogo.Visible = False
        panelSucursal.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelSucursal.Location.X = -576
            panelSucursal.Location = New Point(panelSucursal.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelSucursal.Location.X = -586
            panelSucursal.Location = New Point(panelSucursal.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub picSucursalesListado_Click(sender As Object, e As EventArgs) Handles picSucursalesListado.Click  'abrir form listado sucursales
        picSucursalesGestion.Hide()
        picSucursalesListado.Hide()
        If abierto = False Then
            abierto = True
            panelListadoSucursales.Location = New Point(-586, 0) 'Posicion del panel cuando no está visible
            Do Until panelListadoSucursales.Location.X = 159
                panelListadoSucursales.Location = New Point(panelListadoSucursales.Location.X + 1, 0)
            Loop
            Do Until panelListadoSucursales.Location.X = 169
                panelListadoSucursales.Location = New Point(panelListadoSucursales.Location.X + 1, 0)
                ' Refresh() ' este va
                System.Threading.Thread.Sleep(20)
            Loop
        Else
            MsgBox("Cierre el formulario actual")
        End If
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles PictureBox21.Click  'cerrar form listado sucursales
        abierto = False
        charmLogo.Visible = False
        panelListadoSucursales.Location = New Point(169, 0) 'Posicion del panel cuando si está visible
        Do Until panelListadoSucursales.Location.X = -576
            panelListadoSucursales.Location = New Point(panelListadoSucursales.Location.X - 1, 0)
            ' Refresh() 'Esto soluciona el cuadrado rojo cuando cierra pero va mas lento
        Loop
        Do Until panelListadoSucursales.Location.X = -586
            panelListadoSucursales.Location = New Point(panelListadoSucursales.Location.X - 1, 0)
            ' Refresh() 'este va
            System.Threading.Thread.Sleep(20)
        Loop
        charmLogo.Visible = True
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click  'mostrar ventana config
        ventanaEstadistica.Show()
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click  'añadir sucursal
        'comprobar errores
        conexion()
        Try
            cm.CommandText = "insert into sucursal values (0, " & Chr(34) & txtNomSuc.Text & Chr(34) & "," & Chr(34) & txtCalleSuc.Text & Chr(34) & "," &
                Chr(34) & txtZonaSuc.Text & Chr(34) & ",null," & Chr(34) & txtCiudadSuc.Text & Chr(34) & "," & txtPuertaSuc.Text & "," &
                "'Activo'," & Chr(34) & txtPaisSuc.Text & Chr(34) & ");"
            cm.ExecuteNonQuery()
            MsgBox("Sucursal añadida con exito.")
        Catch
            MsgBox("N°Error 157: Error al añadir sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'refrescar tabla de sucursal
        DataGridSucursal.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM sucursal" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Sucursal")
            DataGridSucursal.DataSource = ds.Tables("Sucursal")
        Catch
            MsgBox("N°Error 158: Error al cargar sucursales.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click  'cargar sucursal
        Dim estado As String
        conexion()
        Try
            cm.CommandText = "SELECT nombre_sucursal,nombre_calle,nombre_zona,nombre_ciudad,numeropuerta,estado,pais FROM sucursal where id_sucursal = " & txtIdSuc.Text & ";"
            read = cm.ExecuteReader
            read.Read()
            txtNomSuc.Text = read.Item(0).ToString
            txtCalleSuc.Text = read.Item(1).ToString
            txtZonaSuc.Text = read.Item(2).ToString
            txtCiudadSuc.Text = read.Item(3).ToString
            txtPuertaSuc.Text = read.Item(4).ToString
            estado = read.Item(5).ToString
            txtPaisSuc.Text = read.Item(6).ToString
        Catch
            MsgBox("N°Error 159: Error al cargar sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        If estado = "Activo" Then
            checkBajaSuc.Checked = False
        Else
            checkBajaSuc.Checked = True
        End If
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click  'modificar suc
        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET nombre_sucursal= " & Chr(34) & txtNomSuc.Text & Chr(34) &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 160: Error al actualizar nombre de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET nombre_calle= " & Chr(34) & txtCalleSuc.Text & Chr(34) &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 161: Error al actualizar calle de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET nombre_zona= " & Chr(34) & txtZonaSuc.Text & Chr(34) &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 162: Error al actualizar zona de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET nombre_ciudad= " & Chr(34) & txtCiudadSuc.Text & Chr(34) &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 163: Error al actualizar ciudad de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET pais = " & Chr(34) & txtPaisSuc.Text & Chr(34) &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 164: Error al actualizar pais de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        conexion()
        Try
            cm.CommandText = "UPDATE sucursal SET numeropuerta= " & txtPuertaSuc.Text &
                " WHERE id_sucursal= " & txtIdSuc.Text & ";"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 165: Error al actualizar numero de puerta de sucursal.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'cambio de estado
        If checkBajaSuc.Checked = True Then 'da de baja a la sucursal
            Dim fechaBaja As String = DateAndTime.Now.ToString("yyyy-MM-dd")
            conexion()
            Try
                cm.CommandText = "UPDATE sucursal SET estado = 'Baja' " &
                    " WHERE id_sucursal= " & txtIdSuc.Text & ";"
                cm.ExecuteNonQuery()
            Catch
                MsgBox("N°Error 166: Error al dar de baja Sucursal (estado).")
            End Try
            cx.Close()
            historial(cm.CommandText)

            'añadir la fecha de baja

            conexion()
            Try
                cm.CommandText = "UPDATE sucursal SET fecha_baja_sucursal =" & Chr(34) & fechaBaja & Chr(34) &
                    " WHERE id_sucursal= " & txtIdSuc.Text & ";"
                cm.ExecuteNonQuery()
                MsgBox("Sucursal dado de baja.")
            Catch
                MsgBox("N°Error 167: Error al dar de baja Sucursal (fecha).")
            End Try
            cx.Close()
            historial(cm.CommandText)

        Else
            'obtiene el estado para ver si se le da de alta 
            Dim estado As String
            conexion()
            Try
                cm.CommandText = "SELECT estado FROM sucursal where id_sucursal = " & txtIdSuc.Text & ";"
                read = cm.ExecuteReader
                read.Read()
                estado = read.Item(0).ToString
            Catch
                MsgBox("N°Error 168: Error al cargar estado de sucursal para dar de alta.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            If estado = "Baja" Then 'dar de alta a la sucursal
                conexion()
                Try
                    cm.CommandText = "UPDATE sucursal SET estado = 'Activo' " &
                        " WHERE id_sucursal= " & txtIdSuc.Text & ";"
                    cm.ExecuteNonQuery()
                    MsgBox("Sucursal dada de alta correctamente.")
                Catch
                    MsgBox("N°Error 169: Error al dar de alta Sucursal (estado).")
                End Try
                cx.Close()
                historial(cm.CommandText)
            End If
        End If

        'refrescar tabla de sucursal
        DataGridSucursal.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM sucursal" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Sucursal")
            DataGridSucursal.DataSource = ds.Tables("Sucursal")
        Catch
            MsgBox("N°Error 170: Error al cargar sucursales.")
        End Try
        cx.Close()

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click  'cargar lsitado caracteristicas
        conexion()
        If DataGridCaracteristica.RowCount > 0 Then
            Me.DataGridCaracteristica.DataSource.clear()
        End If

        Try
            cm.CommandText = "SELECT * FROM tiene,caracteristica where caracteristica.numero_caracteristica = tiene.numero_caracte
                and tiene.padron = " & Chr(34) & buscadorCaracteristica.Text & Chr(34) 'comprobar los nombres de atributos del join
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas Vehiculo")
            DataGridCaracteristica.DataSource = ds.Tables("Caracteristicas Vehiculo")
        Catch
            MsgBox("N°Error 171: Error al cargar caracteristicas.")
        End Try
        cx.Close()
        historial(cm.CommandText)

    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click  'añadir caracteristica general
        conexion()
        Try
            cm.CommandText = "insert into caracteristica values (0, " & Chr(34) & txtNomCaracter.Text & Chr(34) & "," &
               Chr(34) & txtDescCaracter.Text & Chr(34) & ",null," & Chr(34) & "Activo" & Chr(34) & ");"
            cm.ExecuteNonQuery()
            MsgBox("Caracteristica añadida correctamente.")
        Catch
            MsgBox("N°Error 172: Error al añadir caracteristica.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'refrescar las caracteristicas 
        DataGridTodasCaracteristicas.DataSource.clear()
        conexion()
        Try 'carga todas las caracteristicas a su grid
            cm.CommandText = "SELECT * FROM caracteristica where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas")
            DataGridTodasCaracteristicas.DataSource = ds.Tables("Caracteristicas")
        Catch
            MsgBox("N°Error 173: Error al cargar caracteristicas.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click  'cargar caracteristica general
        conexion()

        txtNomCaracter.Text = ""
        txtDescCaracter.Text = ""

        Try
            cm.CommandText = "SELECT nombre_caracteristica,descripcion,estado FROM caracteristica where numero_caracteristica = " & txtNumCaracter.Text & ";"
            read = cm.ExecuteReader
            read.Read()
            txtNomCaracter.Text = read.Item(0).ToString
            txtDescCaracter.Text = read.Item(1).ToString
            'comprueba el estado
            If "Baja" = read.Item(2).ToString Then
                CheckBajaCaracteeristicas.Checked = True
            End If
        Catch
            MsgBox("N°Error 174: Error al cargar caracteristicas.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click  'modificar caracteristica general
        conexion()
        Try
            cm.CommandText = "UPDATE caracteristica SET descripcion= " & Chr(34) & txtDescCaracter.Text & Chr(34) &
                " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
            cm.ExecuteNonQuery()
            MsgBox("Caracteristica actualizada correctamente.")
        Catch
            MsgBox("N°Error 175: Error al actualizar caracteristicas.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        If CheckBajaCaracteeristicas.Checked = True Then 'cambio de estado de caracteristica
            'obtiene la fecha del momento para dar de baja
            Dim fechaBajaC As String
            fechaBajaC = DateAndTime.Now.ToString("yyyy-MM-dd")

            'actualiza el estado, da de baja
            Try
                conexion()

                cm.CommandText = "UPDATE caracteristica SET estado = 'Baja' " &
                    " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
                cm.ExecuteNonQuery()
                historial(cm.CommandText)

                cm.CommandText = "UPDATE caracteristica SET fecha_baja_caracteristica =" & Chr(34) & fechaBajaC & Chr(34) &
                    " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
                cm.ExecuteNonQuery()
                historial(cm.CommandText)

                cx.Close()
            Catch
                MsgBox("N°Error 176: Error al dar de baja caracteristica.")
            End Try
        Else
            'actualiza el estado, da de alta
            Try
                conexion()

                cm.CommandText = "UPDATE caracteristica SET estado = 'Activo' " &
                    " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
                cm.ExecuteNonQuery()
                historial(cm.CommandText)

                cm.CommandText = "UPDATE caracteristica SET fecha_baja_caracteristica = 'null'" & 'fecha de baja nula
                    " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
                cm.ExecuteNonQuery()
                historial(cm.CommandText)

                cx.Close()
            Catch
                MsgBox("N°Error 177: Error al dar de alta caracteristica.")
            End Try
        End If

        'renovar tabla caracteristica
        DataGridTodasCaracteristicas.DataSource.clear()
        conexion()
        Try 'carga todas las caracteristicas a su grid
            cm.CommandText = "SELECT * FROM caracteristica"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas")
            DataGridTodasCaracteristicas.DataSource = ds.Tables("Caracteristicas")
        Catch
            MsgBox("N°Error 178: Error al cargar caracteristicas.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click  'añadir caracteristica a vehiculo
        conexion()
        Try
            cm.CommandText = "insert into tiene values (" & Chr(34) & txtPadronVehiculo.Text & Chr(34) & "," &
                txtNumCaracter2.Text & ");"
            cm.ExecuteNonQuery()
            MsgBox("Caracteristica añadida correctamente al vehiculo.")
        Catch
            MsgBox("N°Error 179: Error al añadir caracteristica al vehiculo.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click  'cargar listado mantenimiento
        Dim hora1, hora2 As DateTime
        Dim fechaIni, fechaFin As String
        Dim erro As Boolean = True

        If DataGridListadoMantenimiento.CurrentCell Is Nothing Then
        Else
            Me.DataGridListadoMantenimiento.DataSource.clear()

        End If
        Try
            hora1 = Date.ParseExact(txtINI.Text, "HH:mm",
           System.Globalization.DateTimeFormatInfo.InvariantInfo)

            hora2 = Date.ParseExact(txtFIN.Text, "HH:mm",
               System.Globalization.DateTimeFormatInfo.InvariantInfo)

            fechaIni = fechaInicioListadoMant.Value.ToString("yyyy-MM-dd") & " " & txtINI.Text
            fechaFin = fechaFinalListadoMant.Value.ToString("yyyy-MM-dd") & " " & txtFIN.Text
            erro = False
        Catch ex As Exception
            MsgBox("Ingrese horas correctas.")
        End Try

        If erro = False Then
            conexion()
            Try
                cm.CommandText = "SELECT * FROM recibe where " & ComboBoxMantenimiento.Text & " = " & Chr(34) & buscadorMantenimiento.Text & Chr(34) &
                    " and fecha_inicial >= " & Chr(34) & fechaIni & Chr(34) & " and fecha_final <= " & Chr(34) & fechaFin & Chr(34)
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "mant")
                DataGridListadoMantenimiento.DataSource = ds.Tables("mant")
            Catch
                MsgBox("N°Error 180: Error al visualizar mantenimientos.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub checkAñadirMant_CheckedChanged(sender As Object, e As EventArgs) Handles checkAñadirMant.CheckedChanged
        If checkAñadirMant.Checked = True Then
            checkModMant.Checked = False
            txtId.Visible = False
            Label56.Visible = False
        End If
    End Sub

    Private Sub checkModMant_CheckedChanged(sender As Object, e As EventArgs) Handles checkModMant.CheckedChanged
        If checkModMant.Checked = True Then
            checkAñadirMant.Checked = False
            txtId.Visible = True
            Label56.Visible = True
        End If
    End Sub

    Private Sub checkAñadirCaract_CheckedChanged(sender As Object, e As EventArgs) Handles checkAñadirCaract.CheckedChanged
        If checkAñadirCaract.Checked = True Then
            checkModCaracte.Checked = False
            txtNumCaracter.Visible = False
            Label62.Visible = False
            CheckBajaCaracteeristicas.Hide()
        End If
    End Sub

    Private Sub checkModCaracte_CheckedChanged(sender As Object, e As EventArgs) Handles checkModCaracte.CheckedChanged
        If checkModCaracte.Checked = True Then
            checkAñadirCaract.Checked = False
            txtNumCaracter.Visible = True
            Label62.Visible = True
            CheckBajaCaracteeristicas.Show()
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click   'añadir mantenimiento
        conexion()
        Try
            cm.CommandText = "insert into mantenimiento values (0, " & Chr(34) & txtBoxMant.Text & Chr(34) & "," & Chr(34) & "Activo" & Chr(34) & " );"
            cm.ExecuteNonQuery()
            MsgBox("Mantenimiento añadido con exito.")
        Catch
            MsgBox("N°Error 181: Error al añadir mantenimiento.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'refrescar los mantenimientos
        DataGridTodosLosMantenimientos.DataSource.clear()
        conexion()
        Try 'carga todas los mantenimientos a su grid
            cm.CommandText = "SELECT * FROM mantenimiento where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Mantenimientos")
            DataGridTodosLosMantenimientos.DataSource = ds.Tables("Mantenimientos")
        Catch
            MsgBox("N°Error 182: Error al cargar mantenimientos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click   'modificar mantenimiento
        conexion()
        Try
            cm.CommandText = "UPDATE mantenimiento SET descripcion= " & Chr(34) & txtBoxMant.Text & Chr(34) &
                " WHERE id_mantenimiento= " & Chr(34) & txtId.Text & Chr(34) & ";"
            cm.ExecuteNonQuery()
            MsgBox("Mantenimiento actualizado con exito.")
        Catch
            MsgBox("N°Error 183: Error al actualizar mantenimiento.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'refrescar los mantenimientos
        DataGridTodosLosMantenimientos.DataSource.clear()
        conexion()
        Try 'carga todas los mantenimientos a su grid
            cm.CommandText = "SELECT * FROM mantenimiento where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Mantenimientos")
            DataGridTodosLosMantenimientos.DataSource = ds.Tables("Mantenimientos")
        Catch
            MsgBox("N°Error 184: Error al cargar mantenimientos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click   'cargar mantenimiento
        conexion()
        Try
            cm.CommandText = "SELECT descripcion FROM mantenimiento where id_mantenimiento  = " & txtId.Text
            read = cm.ExecuteReader
            read.Read()
            txtBoxMant.Text = read.Item(0).ToString
        Catch
            MsgBox("N°Error 185: Error al visualizar mantenimientos.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click  'aplicar mantenimiento a vehiculo
        If txtIdMant.Text = "" Or txtPadron.Text = "" Then
            Label58.BackColor = Color.FromArgb(255, 192, 128)
            Label59.BackColor = Color.FromArgb(255, 192, 128)
        Else
            Label58.BackColor = Color.White
            Label59.BackColor = Color.White

            If CheckApplicarMant.Checked = True Then
                'armar formato de las horas
                Dim fechaINI, inicio, mantenimiento As String
                Dim erro As Boolean = True
                Dim hora1 As DateTime
                fechaINI = fechaIniMant.Value.ToString("yyyy-MM-dd") 'agregar HH:mm
                Try
                    hora1 = Date.ParseExact(txtHoraIni.Text, "HH:mm",
                   System.Globalization.DateTimeFormatInfo.InvariantInfo)

                    inicio = fechaINI & " " & txtHoraIni.Text  'arma el formato de la fecha con la hora

                    erro = False
                Catch ex As Exception
                    MsgBox("N°Error 186: Ingrese horas correctas.")
                End Try

                If erro = False Then
                    'nos fijamos el estado del vehiculo antes del mantenimiento
                    conexion()
                    Try
                        cm.CommandText = "SELECT disponibilidad FROM vehiculo where padron  = " & txtPadron.Text
                        read = cm.ExecuteReader
                        read.Read()
                        mantenimiento = read.Item(0).ToString
                    Catch
                        MsgBox("N°Error 187: Error al obtener estado del vehiculo.")
                    End Try
                    cx.Close()

                    If mantenimiento = "Disponible" Then
                        Try
                            conexion()
                            Try
                                cm.CommandText = "insert into recibe values (" & txtIdMant.Text & "," & Chr(34) & txtPadron.Text & Chr(34) &    'YEAR TO MINUTE
                        "," & Chr(34) & inicio & Chr(34) & ",null);"
                                cm.ExecuteNonQuery()
                            Catch
                                MsgBox("N°Error 188: Error al aplicar mantenimiento al vehiculo.")
                            End Try
                            cx.Close()
                            historial(cm.CommandText)

                            conexion()
                            Try
                                cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Mantenimiento" & Chr(34) &
                        " WHERE padron= " & Chr(34) & txtPadron.Text & Chr(34) & ";"
                                cm.ExecuteNonQuery()
                            Catch
                                MsgBox("N°Error 189: Error al actualizar disponibilidad del vehiculo.")
                            End Try
                            cx.Close()
                            historial(cm.CommandText)

                            MsgBox("Mantenimiento aplicado con exito.")
                        Catch
                            MsgBox("N°Error 190: Error al aplicar Mantenimiento.")
                        End Try
                    Else
                        MsgBox("N°Error 191: No se puede aplicarle un mantenimiento al vehiculo seleccionado.")
                    End If
                End If
            End If

            If CheckFinMant.Checked = True Then
                Dim fechaFin, mantenimiento, inicio, fechaINI As String
                fechaINI = fechaIniMant.Value.ToString("yyyy-MM-dd") 'agregar HH:mm
                fechaFin = DateAndTime.Now.ToString("yyyy-MM-dd hh:mm") 'agregar HH:mm
                inicio = fechaINI & " " & txtHoraIni.Text  'arma el formato de la fecha con la hora
                'nos fijamos el estado del vehiculo antes del mantenimiento
                conexion()
                Try
                    cm.CommandText = "SELECT disponibilidad FROM vehiculo where padron  = " & txtPadron.Text
                    read = cm.ExecuteReader
                    read.Read()
                    mantenimiento = read.Item(0).ToString
                Catch
                    MsgBox("N°Error 192: Error al obtener estado del vehiculo.")
                End Try
                cx.Close()

                If mantenimiento = "Mantenimiento" Then
                    Try
                        conexion()
                        Try
                            'update la hora final
                            cm.CommandText = "UPDATE recibe SET fecha_final = " & Chr(34) & fechaFin & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadron.Text & Chr(34) & " and id_mantenimiento= " & txtIdMant.Text & " and fecha_inicial = " & Chr(34) & inicio & Chr(34)
                            cm.ExecuteNonQuery()
                        Catch
                            MsgBox("N°Error 193: Error al finzalizar mantenimiento al vehiculo.")
                        End Try
                        cx.Close()
                        historial(cm.CommandText)

                        conexion()
                        Try
                            cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Disponible" & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadron.Text & Chr(34) & ";"
                            cm.ExecuteNonQuery()
                        Catch
                            MsgBox("N°Error 194: Error al actualizar disponibilidad del vehiculo.")
                        End Try
                        cx.Close()
                        historial(cm.CommandText)

                        MsgBox("Mantenimiento finzalizado con exito.")
                    Catch
                        MsgBox("N°Error 195: Error al finzalizar Mantenimiento.")
                    End Try
                Else
                    MsgBox("N°Error 196: No se puede finalizar un mantenimiento a un vehiculo Disponible.")
                End If

            End If
        End If
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click   'calculador de precio
        Dim precio As String = 0
        Dim diferencia As Long
        Dim d1 As Date = DateTimePickerFIN.Value
        Dim d2 As Date = DateTimePickerINI.Value

        diferencia = DateDiff("d", d2, d1) 'diferencia de dias entre 2 fechas  

        If diferencia <= 0 Then
            Label118.Text = "N°Error 197: Fechas no validas."
        Else
            conexion()
            Try
                cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & ComboBoxCategoria.Text & Chr(34)
                cm.ExecuteNonQuery()
                read = cm.ExecuteReader
                read.Read()
                precio = read.Item(0).ToString
                Label118.Text = precio * diferencia
                precio = 0
            Catch
                MsgBox("N°Error 198: Error al cargar categorias.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click  'añadir categoria
        conexion()
        Try
            cm.CommandText = "insert into categoria values (" & Chr(34) & txtNombreCategoria.Text & Chr(34) & "," & Chr(34) & txtdescripcionCategoria.Text & Chr(34) & "," &
            txtPrecioCategoria.Text & "," & Chr(34) & "Activo" & Chr(34) & ");"
            cm.ExecuteNonQuery()
            MsgBox("Categoria añadida correctamente.")
        Catch
            MsgBox("N°Error 199: Error al añadir categoria.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'actualizar la tabla en el listado
        DataGridCategorias.DataSource.clear()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategorias.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 200: Error al cargar categorias.")
        End Try
        cx.Close()
    End Sub


    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click  'modificar categoria
        If txtNombreCategoria.Text = "" Then
            MsgBox("N°Error 201: Ingrese el nombre de categoria a modificar.")
        Else
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE categoria SET descripcion_categoria= " & Chr(34) & txtdescripcionCategoria.Text & Chr(34) & 'actualiza la descripcion
                    " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 202: Error al actualizar descripcion de categoria.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE categoria SET precio= " & Chr(34) & txtPrecioCategoria.Text & Chr(34) & 'actualiza el precio
                    " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 203: Error al actualizar precio de categoria.")
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
                        MsgBox("N°Error 204: Error al dar de baja categoria.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                Else 'lo da de alta 
                    conexion()
                    Try
                        cm.CommandText = "UPDATE categoria SET estado = 'Activo'" & 'actualiza el estado, lo da de alta
                        " WHERE nombre_categoria= " & Chr(34) & txtNombreCategoria.Text & Chr(34) & ";"
                        cm.ExecuteNonQuery()
                    Catch
                        MsgBox("N°Error 205: Error al dar de alta categoria.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                End If
                MsgBox("Categoria actualizada correctamente.")

            Catch
                MsgBox("N°Error 206: Error al modificar categoria.")
            End Try
        End If

        'actualizar la tabla en el listado
        DataGridCategorias.DataSource.clear()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM categoria"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategorias.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 207: Error al cargar categorias.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click  'cargar categoria
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
            MsgBox("N°Error 208: Error al cargar categoria.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        If estado = "Baja" Then
            CheckBajaCategoria.Checked = True
        Else
            CheckBajaCategoria.Checked = False
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click   'consulta de vehiculo
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
                MsgBox("N°Error 209: Error al cargar Vehiculos de la sucursal.")
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
                MsgBox("N°Error 210: Error al cargar Vehiculos.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles PictureBox24.Click   'refrescar consulta vehiculo
        If DataGridVehiculo.RowCount > 0 Then
            DataGridVehiculo.DataSource.clear()
        End If

        conexion()
        Try
            cm.CommandText = "SELECT * FROM vehiculo"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 211: Error al cargar Vehiculos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click   'añadir vehiculo

        If txtPadronVehiculo1.Text = "" Or txtAñoVehiculo.Text = "" Or txtMarcaVehiculo.Text = "" Or txtModeloVehiculo.Text = "" Or
           txtMatriculavehiculo.Text = "" Or txtKilometrajeVehiculo.Text = "" Or txtColorVehiculo.Text = "" Then
            txtPadronVehiculo1.BackColor = Color.FromArgb(255, 192, 128)
            txtAñoVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtMarcaVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtModeloVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtMatriculavehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtKilometrajeVehiculo.BackColor = Color.FromArgb(255, 192, 128)
            txtColorVehiculo.BackColor = Color.FromArgb(255, 192, 128)
        Else
            txtPadronVehiculo1.BackColor = Color.White
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
                    cm.CommandText = "insert into vehiculo values (" & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & "," & txtAñoVehiculo.Text & "," &
                   Chr(34) & txtMarcaVehiculo.Text & Chr(34) & "," & Chr(34) & txtModeloVehiculo.Text & Chr(34) & "," & Chr(34) & txtMatriculavehiculo.Text & Chr(34) &
                   "," & Chr(34) & txtKilometrajeVehiculo.Text & Chr(34) & "," & Chr(34) & txtCombustibleVehiculo.Text & Chr(34) & "," & Chr(34) & txtColorVehiculo.Text & Chr(34) & ", 0," &
                   Chr(34) & "Disponible" & Chr(34) & "," & Chr(34) & ComboBoxCategoriaVehiculo.Text & Chr(34) & "," & Chr(34) & "Activo" & Chr(34) & ",null" & ");"
                    cm.ExecuteNonQuery()
                    a = True
                Catch
                    MsgBox("N°Error 212: Error al añadir Vehiculo.")
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
                        MsgBox("N°Error 213: Error al obtener id de sucursal.")
                    End Try
                    cx.Close()

                    conexion()
                    Try
                        'inserta datos a tabla esta, que auto en que sucursal
                        cm.CommandText = "insert into esta values (" & idS & "," & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ");"
                        cm.ExecuteNonQuery()
                        MsgBox("Vehiculo añadido correctamente.")
                    Catch
                        MsgBox("N°Error 214: Error al añadir Vehiculo.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                End If

                'recarga vehiculos de datagrid
                DataGridVehiculo.DataSource.clear()
                conexion()
                Try
                    cm.CommandText = "SELECT * FROM vehiculo"
                    cm.ExecuteNonQuery()
                    da.SelectCommand = cm
                    da.Fill(ds, "Vehiculos")
                    DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
                Catch
                    MsgBox("N°Error 215: Error al cargar Vehiculos.")
                End Try
                cx.Close()
            Else
                MsgBox("N°Error 216: Año del vehiculo no valido.")
            End If
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click   'modificar vehiculo
        Dim fechaBaja As String
        If CheckBoxBaja.Checked Then
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET estado = " & Chr(34) & "Baja" & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                    MsgBox(cm.CommandText)
                Catch
                    MsgBox("N°Error 217: Error al actualizar estado del Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                'añadir fecha de baja
                fechaBaja = DateAndTime.Now.ToString("yyyy-MM-dd")
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET fecha_baja_vehiculo = " & Chr(34) & fechaBaja & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 218: Error al actualizar fecha baja del Vehiculo.")
                End Try
                cx.Close()
                MsgBox("Vehiculo dado de baja correctamente.")
            Catch
                MsgBox("N°Error 219: Error al dar de baja Vehiculo.")
            End Try
            historial(cm.CommandText)
        Else
            Try
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET ano = " & Chr(34) & txtAñoVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 220: Error al actualizar año de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET marca = " & Chr(34) & txtMarcaVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 221: Error al actualizar marca de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET modelo = " & Chr(34) & txtModeloVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 222: Error al actualizar modelo de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET matricula = " & Chr(34) & txtMatriculavehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 223: Error al actualizar matricula de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET nombre_categoria = " & Chr(34) & ComboBoxCategoriaVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 224: Error al actualizar categoria de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET combustible = " & Chr(34) & txtCombustibleVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 225: Error al actualizar combustible de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET kilometraje = " & Chr(34) & txtKilometrajeVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 226: Error al actualizar kilometraje de Vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET color = " & Chr(34) & txtColorVehiculo.Text & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 227: Error al actualizar color de Vehiculo.")
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
                    MsgBox("N°Error 228: Error al obtener id de sucursal.")
                End Try
                cx.Close()

                conexion()
                Try
                    cm.CommandText = "UPDATE esta SET id_sucursal = " & idS &
                    " WHERE padron= " & Chr(34) & txtPadronVehiculo1.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 229: Error al actualizar sucursal del vehiculo.")
                End Try
                cx.Close()

                '--------------------------------------------------------------------

                MsgBox("Vehiculo actualizado correctamente.")
            Catch
                MsgBox("N°Error 230: Error al modificar vehiculo.")
            End Try

        End If

        DataGridVehiculo.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM vehiculo"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 231: Error al cargar vehiculos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click   'cargar vehiculo
        Dim estadoV As String
        Dim a As Boolean = False
        'añadir foto del vehiculo
        Dim fotoo As String

        conexion()
        Try
            cm.CommandText = "SELECT ano,marca,modelo,matricula,combustible,kilometraje,color,nombre_categoria,estado,foto FROM vehiculo where padron = " & Chr(34) & txtPadronVehiculo1.Text & Chr(34)
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
            a = True
        Catch
            MsgBox("N°Error 232: Error al cargar vehiculo.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'añadir foto del vehiculo
        Try
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
        If a = True Then
            Dim idS As Integer
            conexion()
            Try
                cm.CommandText = "SELECT id_sucursal FROM esta where padron = " & Chr(34) & txtPadronVehiculo1.Text & Chr(34)
                read = cm.ExecuteReader
                read.Read()
                idS = read.Item(0)
            Catch
                MsgBox("N°Error 233: Error al cargar id de sucursal.")
            End Try
            cx.Close()

            conexion()
            Try
                cm.CommandText = "SELECT nombre_sucursal FROM sucursal where id_sucursal = " & idS
                read = cm.ExecuteReader
                read.Read()
                ComboBoxSucursalV.SelectedIndex = ComboBoxSucursalV.FindString(read.Item(0).ToString)
            Catch
                MsgBox("N°Error 234: Error al cargar nombre de sucursal.")
            End Try
            cx.Close()
        End If
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
            MsgBox("N°Error 235: Error al cargar cliente.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles PictureBox23.Click  'recargar listado cliente
        DataGridListadoCliente.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM cliente"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Cliente")
            DataGridListadoCliente.DataSource = ds.Tables("Cliente")
        Catch
            MsgBox("N°Error 236: Error al cargar clientes.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click   'añadir cliente

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

            If a >= 18 And a <= 80 Then
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

                        If total Mod 10 + Int32.Parse(cedulac(7)) = 10 Then
                            cedula = True
                        Else
                            MsgBox("N°Error 237: Ingrese una Cedula valida.")
                        End If
                    Catch
                        MsgBox("N°Error 238: Ingrese una Cedula valida.")
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
                        MsgBox("N°Error 239: Error al añadir cliente.")
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
                    MsgBox("N°Error 240: Error al cargar clientes.")
                End Try
                cx.Close()
            Else
                MsgBox("N°Error 241: Cliente debe ser mayor de edad.")
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click   'modificar cliente
        Dim estado As String

        conexion()
        Try
            cm.CommandText = "SELECT estado from cliente where documento_cliente =" & Chr(34) & txtDocumentoCliente.Text & Chr(34) &
                    "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            estado = read.Item(0).ToString
        Catch
            MsgBox("N°Error 242: Error al cargar estado del cliente.")
        End Try
        cx.Close()

        If estado = "Baja" And CheckBoxCliente.Checked = False Then
            conexion()
            Try
                cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Inactivo" & Chr(34) &
            " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34) & ";"
                cm.ExecuteNonQuery()
                MsgBox("Cliente dado de alta correctamente.")
            Catch
                MsgBox("N°Error 243: Error al dar de alta cliente.")
            End Try
            cx.Close()
        End If

        If CheckBoxCliente.Checked Then 'dar de baja al cliente
            If estado = "Activo" Then
                MsgBox("N°Error 244: No se puede dar de baja a un cliente que esté alquilando.")
            Else
                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Baja" & Chr(34) &
                " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & "and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34) & ";"
                    cm.ExecuteNonQuery()
                    MsgBox("Cliente dado de baja correctamente.")
                Catch
                    MsgBox("N°Error 245: Error al dar de baja al Cliente.")
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
                    MsgBox("N°Error 246: Error al actualizar nombre del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET apellido_conductor = " & Chr(34) & txtApellidoCliente.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 247: Error al actualizar apellido del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_ciudad = " & Chr(34) & txtCiudad.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 248: Error al actualizar ciudad del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_calle = " & Chr(34) & txtCalle.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 249: Error al actualizar calle del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET nombre_zona = " & Chr(34) & txtZona.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 250: Error al actualizar zona del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET numeropuerta = " & Chr(34) & txtNumeroPuerta.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 251: Error al actualizar numero de puerta del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET fecha_nacimiento = " & Chr(34) & FechaNacCliente.Value.ToString("yyyy-MM-dd") & Chr(34) & 'fecha
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 252: Error al actualizar fecha de nacimiento del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET telefono = " & Chr(34) & txtTelefonoCliente.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 253: Error al actualizar telefono del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE cliente SET pais = " & Chr(34) & txtPais.Text & Chr(34) &
                    " WHERE documento_cliente= " & Chr(34) & txtDocumentoCliente.Text & Chr(34) & " and tipo_documento = " & Chr(34) & ComboBoxDocumento.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 254: Error al actualizar pais del cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                MsgBox("Cliente actualizado correctamente.")
            Catch
                MsgBox("N°Error 255: Error al modificar Cliente.")
            End Try
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click   'cargar cliente
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
            MsgBox("N°Error 256: Error al cargar Cliente.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click  'listado reserva
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
                MsgBox("N°Error 257: Error al visualizar Reserva.")
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
                MsgBox("N°Error 258: Error al visualizar Reserva.")
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
                MsgBox("N°Error 259: Error al visualizar Reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        ElseIf ComboBoxListadoReserva.Text = "Fecha" Then
            Dim fechaIni, fechaFin As String

            fechaIni = DateTimeIniListadoReserva.Value.ToString("yyyy-MM-dd")
            fechaFin = DateTimeFinListadoReserva.Value.ToString("yyyy-MM-dd")

            conexion()
            Try 'con fechas
                cm.CommandText = "SELECT * FROM reserva,cliente where reserva.fecha_inicio >= " & Chr(34) & fechaIni & Chr(34) & " and reserva.fecha_final <= " & Chr(34) & fechaFin & Chr(34) &
                    " and cliente.documento_cliente = reserva.documento_cliente"
                cm.ExecuteNonQuery()
                da.SelectCommand = cm
                da.Fill(ds, "Reserva")
                DataGridListadoReserva.DataSource = ds.Tables("Reserva")
            Catch
                MsgBox("N°Error 260: Error al visualizar Reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles PictureBox22.Click  'recargar lsitado reserva
        DataGridListadoReserva.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT * FROM reserva"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "rese")
            DataGridListadoReserva.DataSource = ds.Tables("rese")
        Catch
            MsgBox("N°Error 261: Error al cargar Reservas.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxListadoReserva.SelectedIndexChanged  'al cambiar opcion dentro del combobox Listado Reserva
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

    Private Sub btnCheck_Click(sender As Object, e As EventArgs) Handles btnCheck.Click 'check vehiculo en recepcion
        'verifica disponibilidad del vehiculo
        Dim disp, estado As String
        Dim erro As Boolean = True
        conexion()
        Try
            cm.CommandText = "SELECT disponibilidad, estado from vehiculo where padron = " & Chr(34) & txtPadron1.Text & Chr(34) 'comprobar esto!!
            read = cm.ExecuteReader
            read.Read()
            disp = read.Item(0).ToString
            estado = read.Item(1).ToString
            erro = False
        Catch
            MsgBox("N°Error 262: Error al cargar disponibilidad y estado del vehiculo.")
        End Try
        cx.Close()
        If erro = False Then
            If estado = "Activo" And disp = "Disponible" Then
                'obtener nombre sucursal del vehiculo
                'cargar el combobox con el sucursal
                Dim idS As Integer
                conexion()
                Try
                    cm.CommandText = "SELECT id_sucursal from esta where padron = " & Chr(34) & txtPadron1.Text & Chr(34)
                    read = cm.ExecuteReader
                    read.Read()
                    idS = read.Item(0)
                Catch
                    MsgBox("N°Error 263: Error al cargar sucursal del vehiculo.")
                End Try
                cx.Close()

                conexion()
                Try
                    cm.CommandText = "SELECT nombre_sucursal from sucursal where id_sucursal = " & idS
                    read = cm.ExecuteReader
                    read.Read()
                    ComboBoxSucursal.SelectedIndex = ComboBoxSucursal.FindString(read.Item(0).ToString)
                Catch
                    MsgBox("N°Error 264: Error al cargar nombre del sucursal.")
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
                MsgBox("N°Error 265: Vehiculo no disponible.")
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click 'reserrvar
        Dim id, documento, tipoDocumento, padron, FInicio, FFinal, documentoSecundario, tipoDocSecundario, sucursal As String
        Dim diferencia As Integer
        Dim erro As Boolean = False

        documento = txtDocumento.Text
        tipoDocumento = ComboBoxTipoDeDocumento.Text
        padron = txtPadron1.Text
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
                    MsgBox("N°Error 266: Ingrese una Cedula valida.")
                End If
            Catch
                erro = True
                MsgBox("N°Error 267: Ingrese una Cedula valida.")
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
                    MsgBox("N°Error 268: Ingrese una Cedula valida para el segundo conductor.")
                End If
            Catch
                erro = True
                MsgBox("N°Error 269: Ingrese una Cedula valida para el segundo conductor.")
            End Try
        End If

        'comprobar estado del vehiculo
        Dim estadoVehiculo As String
        conexion()
        Try
            cm.CommandText = "SELECT disponibilidad FROM vehiculo where padron = " & Chr(34) & txtPadron1.Text & Chr(34)
            cm.ExecuteNonQuery()
            read = cm.ExecuteReader
            read.Read()
            estadoVehiculo = read.Item(0).ToString
        Catch
            MsgBox("N°Error 270: Error al cargar estado del Vehiculo.")
        End Try
        cx.Close()

        If Not estadoVehiculo = "Disponible" Then
            erro = True
            MsgBox("N°Error 271: Vehiculo no disponible.")
        End If

        'comprobar fechas 
        If DateTimeFechaInicioReserva.Value < DateTime.Today Then
            erro = True
            MsgBox("N°Error 272: Ingrese fechas correctas.")
        Else
            If DateTimeFechaInicioReserva.Value > DateTimeFechaFinReserva.Value Then
                erro = True
                MsgBox("N°Error 273: Ingrese fechas correctas.")
            End If
        End If

        'comprobar campos vacios

        If ComboBoxTipoDeDocumento.Text = "" Or txtDocumento.Text = "" Or txtPadron1.Text = "" Then
            erro = True
            Label7.BackColor = Color.FromArgb(255, 192, 128)
            Label6.BackColor = Color.FromArgb(255, 192, 128)
            Label111.BackColor = Color.FromArgb(255, 192, 128)
            MsgBox("N°Error 274: Llene los campos obligatorios.")
        End If

        If CheckBoxConsuctorSec.Checked = True Then
            If ComboBoxTipoDoc2do.Text = "" Or txtDoc2do.Text = "" Then
                erro = True
                LabelDoc2.BackColor = Color.FromArgb(255, 192, 128)
                labelDoc2a.BackColor = Color.FromArgb(255, 192, 128)
                MsgBox("N°Error 275: Llene los campos obligatorios.")
            End If
        End If

        'cambiar la sucursal del vehiculo - traslado
        conexion()

        cx.Close()

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 'cargar reserva
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
            txtPadron1.Text = read.Item(2).ToString
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
            MsgBox("N°Error 276: Error al cargar reserva.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        If erro = False Then


            'cargar la sucursal del vehiculo
            Dim ids As Integer
            Dim erro2 As Boolean = True
            conexion()
            Try
                cm.CommandText = "SELECT id_sucursal from esta where padron = " & Chr(34) & txtPadron1.Text & Chr(34)
                read = cm.ExecuteReader
                read.Read()
                ids = read.Item(0)
                erro2 = False
            Catch
                MsgBox("N°Error 277: Error al obtener id de sucursal.")
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
                    MsgBox("N°Error 278: Error al obtener nombre de sucursal.")
                End Try
                cx.Close()
            End If
        End If
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click 'modificar reserva
        'depende del estaod de la reserva si se puede modificar o no
        Dim estado, padronAnterior As String

        conexion()
        Try
            cm.CommandText = "SELECT estado_reserva FROM reserva where id_reserva = " & txtIdReserva.Text
            read = cm.ExecuteReader
            read.Read()
            estado = read.Item(0).ToString
        Catch
            MsgBox("N°Error 279: Error al obtener estado de reserva.")
        End Try
        cx.Close()
        historial(cm.CommandText)
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
                    MsgBox("N°Error 280: Error al obtener vehiculo de reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Disponible" & Chr(34) &
                    " WHERE padron= " & Chr(34) & padronAnterior & Chr(34) 'cambiar estado del anterior vehiculo a disponible 
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 281: Error al actualizar disponibilidad del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                '------------------------------------------------------------------------------------
                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET padron = " & Chr(34) & txtPadron.Text & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 282: Error al actualizar vehiculo de la reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Reservado" & Chr(34) &
                    " WHERE padron= " & txtPadron.Text 'cambiar estado del nuevo a Reservado
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 283: Error al actualizar disponibilidad del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET fecha_inicio = " & Chr(34) & DateTimeFechaInicioReserva.Value.ToString("yyyy-MM-dd") & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 284: Error al actualizar fecha inicial de la reserva.")
                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE reserva SET fecha_final = " & Chr(34) & DateTimeFechaFinReserva.Value.ToString("yyyy-MM-dd") & Chr(34) &
                    " WHERE id_reserva= " & txtIdReserva.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 285: Error al actualizar fecha final de la reserva.")
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
                    MsgBox("N°Error 286: Error al obtener id de sucursal.")
                End Try
                cx.Close()

                conexion()
                Try
                    cm.CommandText = "UPDATE esta SET id_sucursal = " & idS &
                    " WHERE padron= " & txtPadron.Text
                    cm.ExecuteNonQuery()
                Catch
                    MsgBox("N°Error 287: Error al actualizar sucursal del vehiculo.")
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
                        MsgBox("N°Error 288: Error al actualizar documento segundo conductor de la reserva.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)

                    conexion()
                    Try
                        cm.CommandText = "UPDATE reserva SET tipo_documento2do = " & Chr(34) & ComboBoxTipoDoc2do.Text & Chr(34) &
                        " WHERE id_reserva= " & txtIdReserva.Text
                        cm.ExecuteNonQuery()
                    Catch
                        MsgBox("N°Error 289: Error al actualizar tipo de documento segundo conductor de la reserva.")
                    End Try
                    cx.Close()
                    historial(cm.CommandText)
                End If

                MsgBox("Reserva actualizada correctamente.")
            Catch
                MsgBox("N°Error 290: Error al modificar Reserva.")
            End Try

        Else
            MsgBox("N°Error 291: No se puede modificar un Alquiler.")
        End If
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click 'alquilar
        Dim idRes, FRetiro, sucursal As String
        Dim diferencia As Integer
        Dim erro As Boolean = True
        'fecha retiro
        FRetiro = DateTimePicker4.Value.ToString("yyyy-MM-dd")

        'nombre de sucursal
        sucursal = ComboBoxSucursal.Text

        'diferencia de dias de fechas
        diferencia = DateDiff("d", DateTimeFechaInicioReserva.Value, DateTimeFechaFinReserva.Value)

        'fijar si la fecha de aluiler esta entre las de inicio y final
        If DateTimeFechaInicioReserva.Value < DateTimePicker4.Value And DateTimeFechaFinReserva.Value > DateTimePicker4.Value Then
            erro = False
        Else
            MsgBox("N°Error 292: Error Con fecha de retiro.")
        End If

        If erro = False Then
            idRes = txtIdReserva.Text
            BoletaReserva.generarBoletaAlquiler(idRes, diferencia, FRetiro, sucursal) 'generar boleta
            BoletaReserva.Show()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'finalizar alquiler
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

        'diferencia de dias de fechas
        diferencia = DateDiff("d", DateTimeFechaInicioReserva.Value, DateTimeFechaFinReserva.Value)

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

        If erro = False Then
            idRes = txtIdReserva.Text
            'abrir boleta fin alquiler  
            BoletaReserva.generarBoletaFinAlquiler(idRes, diferencia, FFinal, sucursal)
            BoletaReserva.Show()
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

    Private Sub txtTelefonoCliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTelefonoCliente.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumeroPuerta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNumeroPuerta.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumCaracter_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumCaracter2_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtId_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtIdMant_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtIdSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtIdSuc.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPuertaSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPuertaSuc.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtIdReserva_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtIdReserva.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPais_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPais.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtCiudad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCiudad.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtZona_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtZona.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtCalle_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCalle.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtApellidoCliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtApellidoCliente.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtNombreCliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNombreCliente.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtColorVehiculo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtColorVehiculo.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtNombreCategoria_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNombreCategoria.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtNomSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNomSuc.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtCiudadSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCiudadSuc.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtZonaSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtZonaSuc.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtCalleSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCalleSuc.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtNomCaracter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNomCaracter.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtNumCaracter_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles txtNumCaracter.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumCaracter2_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles txtNumCaracter2.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtIdMant_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles txtIdMant.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtId_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles txtId.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPaisSuc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPaisSuc.KeyPress
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub CheckApplicarMant_CheckedChanged(sender As Object, e As EventArgs) Handles CheckApplicarMant.CheckedChanged
        If CheckApplicarMant.Checked = True Then
            CheckFinMant.Checked = False
            txtHoraIni.Show()
            fechaIniMant.Show()
            Label53.Show()
            Label60.Show()
        End If
    End Sub

    Private Sub CheckFinMant_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFinMant.CheckedChanged
        If CheckFinMant.Checked = True Then
            CheckApplicarMant.Checked = False
            txtHoraIni.Show()
            fechaIniMant.Show()
            Label53.Show()
            Label60.Show()
        End If
    End Sub
End Class
