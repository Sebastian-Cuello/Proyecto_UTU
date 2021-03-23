Imports System.Data.Odbc
'Todo hace algo, deberia andar 
Public Class ventanaMecanico
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet
    Dim read As OdbcDataReader
    Dim abierto As Boolean = False 'Si hay un formulario abierto será verdadero
    Dim caso6 As Boolean = False
    Dim caso5 As Boolean = False
    Dim opcion As Integer 'opcion del menu para desplegar submenus
    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=" & LoginSU.user & ";PWD=" & LoginSU.contraseña
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 17: No hay conexion con la base de datos.")
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
            MsgBox("N°Error 18: Error al mantener historial.")
        End Try
        cx.Close()
    End Sub

    Private Sub ventanaMecanico_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = False

        conexion()
        Try 'carga todas las caracteristicas a su grid
            cm.CommandText = "SELECT * FROM caracteristica where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas")
            DataGridTodasCaracteristicas.DataSource = ds.Tables("Caracteristicas")
        Catch
            MsgBox("N°Error 19: Error al cargar caracteristicas.")
        End Try
        cx.Close()

        conexion()
        Try 'carga todas los mantenimientos a su grid
            cm.CommandText = "SELECT * FROM mantenimiento where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Mantenimientos")
            DataGridTodosLosMantenimientos.DataSource = ds.Tables("Mantenimientos")
        Catch
            MsgBox("N°Error 20: Error al cargar mantenimientos.")
        End Try
        cx.Close()

    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs)  'Cierra toda la aplicacion
        Application.Exit()
    End Sub

    Private Sub pictureCerrarSesion_Click(sender As Object, e As EventArgs) Handles pictureCerrarSesion.Click 'flecha cerrar sesion
        Me.Close()
        LoginSU.Show()
    End Sub

    '------------------------------Botones En Imagenes---------------------------------------------------------
    Private Sub botonesOpciones() 'Abre el boton correcto de la opcion y esconde los otros
        Select Case opcion

            Case 5
                If caso5 = False Then
                    picMantenimientoGestion.Show()
                    picMantenimientoListado.Show()
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
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
                    picMantenimientoGestion.Hide()
                    picMantenimientoListado.Hide()
                    caso6 = True
                Else
                    picCaracteristicasGestion.Hide()
                    picCaracteristicasListado.Hide()
                    caso6 = False
                End If

        End Select
    End Sub

    Private Sub btnMantenimiento_Click(sender As Object, e As EventArgs) Handles btnMantenimiento.Click    'click mantenimiento
        opcion = 5
        botonesOpciones()
    End Sub

    Private Sub picMantenimientoGestion_Click(sender As Object, e As EventArgs) Handles picMantenimientoGestion.Click    'abrir form mantenimiento
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

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click 'cerrar form manenimiento
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

    Private Sub picMantenimientoListado_Click(sender As Object, e As EventArgs) Handles picMantenimientoListado.Click    'Abrir form Listado Mantenimiento
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

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click 'cerrar form listado mantenimiento
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

    Private Sub btnCaracteristica_Click(sender As Object, e As EventArgs) Handles btnCaracteristica.Click    'click caracteristicas
        opcion = 6
        botonesOpciones()
    End Sub

    Private Sub picCaracteristicasGestion_Click(sender As Object, e As EventArgs) Handles picCaracteristicasGestion.Click    'abrir form caracteristica
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

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click 'Cerrar form caracteristicas
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

    Private Sub picCaracteristicasListado_Click(sender As Object, e As EventArgs) Handles picCaracteristicasListado.Click    'abrir form listado caracteristicas
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

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click 'cerrar form listado caracteristicas
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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)  'mostrar ventana config
        ventanaEstadistica.Show()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click  'cargar listado mantenimientos
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
            fechaFin = fechaFinalListadoMant.Value.ToString("yyyy-MM-dd ") & " " & txtFIN.Text
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
                MsgBox("N°Error 21: Error al visualizar mantenimientos.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click  'cargar listado caracteristicas
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
            MsgBox("N°Error 22: Error al visualizar caracteristicas.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click  'aplicar mantenimiento 
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
                    MsgBox("Ingrese horas correctas.")
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
                        MsgBox("N°Error 23: Error al obtener estado del vehiculo.")
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
                                MsgBox("N°Error 24: Error al aplicar mantenimiento al vehiculo.")
                            End Try
                            cx.Close()
                            historial(cm.CommandText)

                            conexion()
                            Try
                                cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Mantenimiento" & Chr(34) &
                        " WHERE padron= " & Chr(34) & txtPadron.Text & Chr(34) & ";"
                                cm.ExecuteNonQuery()
                            Catch
                                MsgBox("N°Error 25: Error al actualizar disponibilidad del vehiculo.")
                            End Try
                            cx.Close()
                            historial(cm.CommandText)

                            MsgBox("Mantenimiento aplicado con exito.")
                        Catch
                            MsgBox("N°Error 26: Error al aplicar Mantenimiento.")
                        End Try
                    Else
                        MsgBox("N°Error 27: No se puede aplicarle un mantenimiento al vehiculo seleccionado.")
                    End If
                End If
            End If

            If CheckFinMant.Checked = True Then
                Dim fechaFin, mantenimiento, fechaINI, inicio As String
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
                    MsgBox("N°Error 28: Error al obtener estado del vehiculo.")
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
                            MsgBox("N°Error 29: Error al finzalizar mantenimiento al vehiculo.")
                        End Try
                        cx.Close()
                        historial(cm.CommandText)

                        conexion()
                        Try
                            cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Disponible" & Chr(34) &
                    " WHERE padron= " & Chr(34) & txtPadron.Text & Chr(34) & ";"
                            cm.ExecuteNonQuery()
                        Catch
                            MsgBox("N°Error 30: Error al actualizar disponibilidad del vehiculo.")
                        End Try
                        cx.Close()
                        historial(cm.CommandText)

                        MsgBox("Mantenimiento finzalizado con exito.")
                    Catch
                        MsgBox("N°Error 31: Error al finzalizar Mantenimiento.")
                    End Try
                Else
                    MsgBox("N°Error 32: No se puede finalizar un mantenimiento a un vehiculo Disponible.")
                End If

            End If
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click  'añadir mantenimiento general
        conexion()
        Try
            cm.CommandText = "insert into mantenimiento values (0, " & Chr(34) & txtBoxMant.Text & Chr(34) & "," & Chr(34) & "Activo" & Chr(34) & " );"
            cm.ExecuteNonQuery()
            MsgBox("Mantenimiento añadido con exito.")
        Catch
            MsgBox("N°Error 33: Error al añadir mantenimiento.")
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
            MsgBox("N°Error 34: Error al cargar mantenimientos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click  'modificar un mantenimiento
        conexion()
        Try
            cm.CommandText = "UPDATE mantenimiento SET descripcion= " & Chr(34) & txtBoxMant.Text & Chr(34) &
                " WHERE id_mantenimiento= " & Chr(34) & txtId.Text & Chr(34) & ";"
            cm.ExecuteNonQuery()
            MsgBox("Mantenimiento actualizado con exito.")
        Catch
            MsgBox("N°Error 35: Error al actualizar mantenimientos.")
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
            MsgBox("N°Error 36: Error al cargar mantenimientos.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click  'cargar el mantenimiento
        If txtId.Text = "" Then
            Label56.BackColor = Color.FromArgb(255, 192, 128)
        Else
            Label56.BackColor = Color.White
            conexion()
            Try
                cm.CommandText = "SELECT descripcion FROM mantenimiento where id_mantenimiento  = " & txtId.Text
                read = cm.ExecuteReader
                read.Read()
                txtBoxMant.Text = read.Item(0).ToString
            Catch
                MsgBox("N°Error 37: Error al visualizar mantenimientos.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click  'añadir caracteristica general
        conexion()
        Try
            cm.CommandText = "insert into caracteristica values (0, " & Chr(34) & txtNomCaracter.Text & Chr(34) & "," &
               Chr(34) & txtDescCaracter.Text & Chr(34) & ",null," & Chr(34) & "Activo" & Chr(34) & ");"
            cm.ExecuteNonQuery()
            MsgBox("Caracteristica añadida correctamente.")
        Catch
            MsgBox("N°Error 38: Error al añadir caracteristica.")
        End Try
        cx.Close()
        historial(cm.CommandText)

        'refrescar las caracteristicas
        DataGridTodasCaracteristicas.DataSource.clea()
        conexion()
        Try 'carga todas las caracteristicas a su grid
            cm.CommandText = "SELECT * FROM caracteristica where estado = 'Activo'"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Caracteristicas")
            DataGridTodasCaracteristicas.DataSource = ds.Tables("Caracteristicas")
        Catch
            MsgBox("N°Error 39: Error al cargar caracteristica.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click  'cargar  caracteristica
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
            MsgBox("N°Error 40: Error al cargar caracteristica.")
        End Try
        historial(cm.CommandText)
        cx.Close()
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click  'moodificar caracteristica
        If txtDescCaracter.Text = "" And txtNomCaracter.Text = "" Then
            Label66.BackColor = Color.FromArgb(255, 192, 128)
            Label65.BackColor = Color.FromArgb(255, 192, 128)
        Else
            Label66.BackColor = Color.White
            Label65.BackColor = Color.White
            conexion()
            Try
                cm.CommandText = "UPDATE caracteristica SET descripcion= " & Chr(34) & txtDescCaracter.Text & Chr(34) &
                    " WHERE numero_caracteristica= " & txtNumCaracter.Text & ";"
                cm.ExecuteNonQuery()
                MsgBox("Caracteristica actualizada correctamente.")
            Catch
                MsgBox("N°Error 41: Error al actualizar caracteristica.")
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
                    MsgBox("N°Error 42: Error al dar de baja la caracteristica.")
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
                    MsgBox("N°Error 43: Error al dar de alta la caracteristica.")
                End Try
            End If
        End If

    End Sub



    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click  'añadir caracteristica a vehiculo
        If txtNumCaracter2.Text = "" And txtPadronVehiculo.Text = "" Then
            Label69.BackColor = Color.FromArgb(255, 192, 128)
            Label70.BackColor = Color.FromArgb(255, 192, 128)
        Else
            Label69.BackColor = Color.White
            Label70.BackColor = Color.White
            conexion()
            Try
                cm.CommandText = "insert into tiene values (" & Chr(34) & txtPadronVehiculo.Text & Chr(34) & "," &
                    txtNumCaracter2.Text & ");"
                cm.ExecuteNonQuery()
                MsgBox("Caracteristica añadida correctamente al vehiculo.")
            Catch
                MsgBox("N°Error 44: Error al añadir caracteristica.")
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

    Private Sub txtNumCaracter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNumCaracter.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNumCaracter2_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNumCaracter2.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtIdMant_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtIdMant.KeyPress, txtId.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtNomCaracter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNomCaracter.KeyPress
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
            Label7.Show()
            Label60.Show()
        End If

    End Sub

    Private Sub CheckFinMant_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFinMant.CheckedChanged
        If CheckFinMant.Checked = True Then
            CheckApplicarMant.Checked = False
            txtHoraIni.Show()
            fechaIniMant.Show()
            Label7.Show()
            Label60.Show()
        End If
    End Sub
End Class