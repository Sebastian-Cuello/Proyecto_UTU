Imports System.Data.Odbc
Public Class Invitado
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim cm2 As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet
    Dim read As OdbcDataReader
    Dim abierto As Boolean = False 'Si hay un formulario abierto será verdadero

    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=cuelloa;PWD=setododetaller"
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 7: No hay conexion con la base de datos.")
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
            cm.CommandText = "insert into historial values (0," & Chr(34) & LocalHostaddress & Chr(34) & "," & Chr(34) & "cuelloa" & Chr(34) & "," & Chr(34) &
            arra & Chr(34) & "," & Chr(34) & hora & Chr(34) & ");"
            cm.ExecuteNonQuery()
        Catch
            MsgBox("N°Error 8: Error al mantener historial.")
            MsgBox(cm.CommandText)
        End Try
        cx.Close()
    End Sub

    Private Sub ventanaInvitado_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = False
        Dim lineas As Integer = 0

        conexion()
        Try 'carga combobox categoria
            cm.CommandText = "SELECT nombre_categoria FROM categoria where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds)
            ComboBoxCategoria.DataSource = ds.Tables(0)
            ComboBoxCategoria.DisplayMember = ds.Tables(0).Columns(0).Caption.ToString
        Catch
            MsgBox("N°Error 9: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT ano,marca,modelo,kilometraje,combustible,color,nombre_categoria FROM vehiculo where disponibilidad = 'Disponible'" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 10: Error al cargar vehiculos.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT nombre_categoria, descripcion_categoria,precio FROM categoria where estado = 'Activo'" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Categoria")
            DataGridCategorias.DataSource = ds.Tables("Categoria")
        Catch
            MsgBox("N°Error 11: Error al cargar categorias.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT * FROM sucursal" 'llena datagrid por primera vez al iniciar
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Sucursal")
            DataGridSucursales.DataSource = ds.Tables("Sucursal")
        Catch
            MsgBox("N°Error 12: Error al cargar sucursales.")
        End Try
        cx.Close()
    End Sub

    Private Sub pictureCerrarSesion_Click(sender As Object, e As EventArgs) Handles pictureCerrarSesion.Click
        Me.Close()
        LoginSU.Show()
    End Sub

    Private Sub btnVehiculo_Click(sender As Object, e As EventArgs) Handles btnVehiculo.Click 'abrir listado vehiculo
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

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click 'cerrar form listado vehiculo
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

    Private Sub btnCategoria_Click(sender As Object, e As EventArgs) Handles btnCategoria.Click 'abrir listado categoria
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

    Private Sub btnSucursales_Click(sender As Object, e As EventArgs) Handles btnSucursales.Click 'abrir listado sucursales
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

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click 'cerrar form listado sucursales
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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)  'boton tuerquita
        ventanaEstadistica.Show() 'abrir ventana config
    End Sub

    Private Sub ImagenVehiculo_Click(sender As Object, e As EventArgs) Handles ImagenVehiculo.Click 'Actualizar tabla seccion vehiculo
        Me.DataGridVehiculo.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT ano,marca,modelo,kilometraje,combustible,color,nombre_categoria FROM vehiculo where " & ComboBoxVehiculo.Text & " = " & Chr(34) & BuscadorVehiculo.Text & Chr(34) 'Chr(34) COMILLA
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 13: Error al visualizar vehiculos.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles  PictureBox13.Click 'refrescar consulta de vehiculo
        Me.DataGridVehiculo.DataSource.clear()
        conexion()
        Try
            cm.CommandText = "SELECT ano,marca,modelo,kilometraje,combustible,color,nombre_categoria FROM vehiculo where disponibilidad = 'Disponible'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridVehiculo.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 14: Error al cargar vehiculos.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'calcular precio 
        Dim precio As String = 0
        Dim diferencia As Long
        Dim d1 As Date = DateTimePickerFIN.Value
        Dim d2 As Date = DateTimePickerINI.Value

        diferencia = DateDiff("d", d2, d1) 'diferencia de dias entre 2 fechas  

        If diferencia <= 0 Then
            Label7.Text = "N°Error 15: Fechas no validas."
        Else
            conexion()
            Try
                cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & ComboBoxCategoria.Text & Chr(34)
                cm.ExecuteNonQuery()
                read = cm.ExecuteReader
                read.Read()
                precio = read.Item(0).ToString
                Label7.Text = precio * diferencia
                precio = 0
            Catch
                MsgBox("N°Error 16: Error al obtener precio de categoria.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub
End Class