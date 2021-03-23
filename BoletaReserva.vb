Imports System.Data.Odbc
Public Class BoletaReserva
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet
    Dim read As OdbcDataReader
    Dim estado As String

    Private Sub conexion()
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=" & LoginSU.user & ";PWD=" & LoginSU.contraseña
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 294: No hay conexion con la base de datos.")
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
            MsgBox("N°Error 295: Error al mantener historial.")
        End Try
        cx.Close()
    End Sub

    Private Sub BoletaReserva_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = False
    End Sub

    Public Sub generarBoleta(documento As String, tipoDocumento As String, padron As String,
                                  FInicio As String, FFinal As String, documentoSecundario As String,
                                  tipoDocSecundario As String, diferencia As Integer, sucursal As String)
        estado = "Reserva"
        Dim Categoria As String
        Dim idRes As Integer
        'obtener nombre y apellido conductor
        conexion()
        Try
            cm.CommandText = "SELECT nombre_conductor,apellido_conductor FROM cliente where documento_cliente = " & Chr(34) & documento & Chr(34) &
            "and tipo_documento = " & Chr(34) & tipoDocumento & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblNombre.Text = read.Item(0).ToString
            lblApellido.Text = read.Item(1).ToString
        Catch
            MsgBox("N°Error 296: Error al cargar nombre y apellido del cliente.")
        End Try
        cx.Close()

        'obtener id de la reserva
        conexion()
        Try
            cm.CommandText = "SELECT count(*) from reserva"
            read = cm.ExecuteReader
            read.Read()
            idRes = read.Item(0)
        Catch
            MsgBox("N°Error 297: Error al generar id de reserva.")
        End Try
        cx.Close()
        idRes = idRes + 1
        Label3.Text = idRes 'id de la reserva

        'obtener precio categoria-------------------------------------------------------------------------------
        conexion()
        Try
            cm.CommandText = "SELECT nombre_categoria FROM vehiculo where padron = " & Chr(34) & padron & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Categoria = read.Item(0).ToString
        Catch
            MsgBox("N°Error 298: Error al cargar el nombre de categoria.")
        End Try
        cx.Close()

        Dim valor As Integer
        conexion()
        Try
            cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & Categoria & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Label13.Text = read.Item(0).ToString * diferencia 'calculo de precio
            valor = read.Item(0) * diferencia * 0.25 'calculo de deposito 25% del precio
            LabelDepo.Text = valor
        Catch
            MsgBox("N°Error 299: Error al cargar el precio de la categoria.")
        End Try
        cx.Close()

        'obtener datos de la sucursal
        conexion()
        Try
            cm.CommandText = "SELECT nombre_ciudad,nombre_zona,nombre_calle,numeropuerta FROM sucursal where nombre_sucursal = " & Chr(34) & sucursal & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblCiudad.Text = read.Item(0).ToString
            lblZona.Text = read.Item(1).ToString
            lblCalle.Text = read.Item(2).ToString
            lblPuerta.Text = read.Item(3).ToString
        Catch
            MsgBox("N°Error 300: Error al obtener datos de sucursal.")
        End Try
        cx.Close()

        If documentoSecundario = "" Then
            Label1.Hide()
            lblTDocSec.Hide()
            lblDocSec.Hide()
        End If

        lblTipoDoc.Text = tipoDocumento
        lblDoc.Text = documento
        lblFInicio.Text = FInicio
        lblFFin.Text = FFinal

        lblDocSec.Text = documentoSecundario
        lblTDocSec.Text = tipoDocSecundario
        lblPadron.Text = padron
        'andando

        'armar la boleta con los datos y hacer la insercion en el boton de aceptar
    End Sub

    Public Sub generarBoletaAlquiler(idRes As String, diferencia As Integer, FRetiro As String, sucursal As String) 'generar boleta alquiler
        Label15.Text = "Boleta de Alquiler"
        Button1.Text = "Alquilar"
        estado = "Alquiler"

        'obtener todos los datos de la reserva para llenar y seleccionar la nueva fecha 
        Dim d1, d2 As Date
        conexion()
        Try
            cm.CommandText = "SELECT id_reserva,documento_cliente,tipo_documento,padron,fecha_inicio,fecha_final" &
                ", documento_2doconductor, tipo_documento2do,deposito FROM reserva where id_reserva = " & idRes
            read = cm.ExecuteReader
            read.Read()
            Label3.Text = read.Item(0).ToString
            lblDoc.Text = read.Item(1).ToString
            lblTipoDoc.Text = read.Item(2).ToString
            lblPadron.Text = read.Item(3).ToString
            lblFInicio.Text = read.Item(4).ToString.Split(" "c)(0)
            lblFFin.Text = read.Item(5).ToString.Split(" "c)(0)
            lblDocSec.Text = read.Item(6).ToString
            lblTDocSec.Text = read.Item(7).ToString
            LabelDepo.Text = read.Item(8).ToString
        Catch
            MsgBox("N°Error 301: Error al obtener datos de reserva.")
        End Try
        cx.Close()

        'obtener datos de la sucursal
        conexion()
        Try
            cm.CommandText = "SELECT nombre_ciudad,nombre_zona,nombre_calle,numeropuerta FROM sucursal where nombre_sucursal = " & Chr(34) & sucursal & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblCiudad.Text = read.Item(0).ToString
            lblZona.Text = read.Item(1).ToString
            lblCalle.Text = read.Item(2).ToString
            lblPuerta.Text = read.Item(3).ToString
        Catch
            MsgBox("N°Error 302: Error al obtener datos de sucursal.")
        End Try
        cx.Close()

        'ocultar label de datos para el segundo conductor 
        If Not lblDocSec.Text = "" Then
            Label1.Show()
            lblTDocSec.Show()
            lblDocSec.Show()
        End If

        'obtener nombre y apellido cliente
        Dim Categoria As String
        conexion()
        Try
            cm.CommandText = "SELECT nombre_conductor,apellido_conductor FROM cliente where documento_cliente = " & Chr(34) & lblDoc.Text & Chr(34) &
            "and tipo_documento = " & Chr(34) & lblTipoDoc.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblNombre.Text = read.Item(0).ToString
            lblApellido.Text = read.Item(1).ToString
        Catch
            MsgBox("N°Error 303: Error al cargar nombre y apellido del cliente.")
        End Try
        cx.Close()

        'obtener precio categoria-------------------------------------------------------------------------------
        conexion()
        Try
            cm.CommandText = "SELECT nombre_categoria FROM vehiculo where padron = " & Chr(34) & lblPadron.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Categoria = read.Item(0).ToString
        Catch
            MsgBox("N°Error 304: Error al cargar nombre de categoria.")
        End Try
        cx.Close()

        conexion()
        Try
            cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & Categoria & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Label13.Text = read.Item(0).ToString * diferencia 'calculo de precio
            LabelDepo.Text = read.Item(0).ToString * diferencia * 0.25 'calculo de deposito 25% del precio
        Catch
            MsgBox("N°Error 305: Error al cargar precio de categoria.")
        End Try
        cx.Close()

        Label16.Show()
        Label14.Show()
        Label14.Text = FRetiro

    End Sub



    Public Sub generarBoletaFinAlquiler(idRes As String, diferencia As Integer, FFinal As String, sucursal As String) 'generar boleta finalizar alquiler 
        estado = "Fin Alquiler"
        Label15.Text = "Boleta de Devolucion"
        Button1.Text = "Finalizar Alquiler"
        Label16.Show()
        Label14.Show()
        Label17.Show()
        Label18.Show()

        'obtener todos los datos del alquiler para llenar y seleccionar la fecha de devolucion 
        conexion()
        Try
            cm.CommandText = "SELECT id_reserva,documento_cliente,tipo_documento,padron,fecha_inicio,fecha_final" &
                ", documento_2doconductor, tipo_documento2do,deposito,fecha_retiro FROM reserva where id_reserva = " & idRes
            read = cm.ExecuteReader
            read.Read()
            Label3.Text = read.Item(0).ToString
            lblDoc.Text = read.Item(1).ToString
            lblTipoDoc.Text = read.Item(2).ToString
            lblPadron.Text = read.Item(3).ToString

            lblFInicio.Text = read.Item(4).ToString.Split(" "c)(0)
            lblFFin.Text = read.Item(5).ToString.Split(" "c)(0)
            lblDocSec.Text = read.Item(6).ToString
            lblTDocSec.Text = read.Item(7).ToString
            LabelDepo.Text = read.Item(8).ToString
            Label14.Text = read.Item(9).ToString().Split(" "c)(0)
        Catch
            MsgBox("N°Error 306: Error al obtener datos del alquiler.")
        End Try
        cx.Close()

        Label18.Text = FFinal

        'obtener datos de la sucursal
        conexion()
        Try
            cm.CommandText = "SELECT nombre_ciudad,nombre_zona,nombre_calle,numeropuerta FROM sucursal where nombre_sucursal = " & Chr(34) & sucursal & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblCiudad.Text = read.Item(0).ToString
            lblZona.Text = read.Item(1).ToString
            lblCalle.Text = read.Item(2).ToString
            lblPuerta.Text = read.Item(3).ToString
        Catch
            MsgBox("N°Error 307: Error al obtener datos de sucursal.")
        End Try
        cx.Close()

        If Not lblDocSec.Text = "" Then
            Label1.Show()
            lblTDocSec.Show()
            lblDocSec.Show()
        End If

        'obtener nombre y apellido cliente
        Dim Categoria As String
        conexion()
        Try
            cm.CommandText = "SELECT nombre_conductor,apellido_conductor FROM cliente where documento_cliente = " & Chr(34) & lblDoc.Text & Chr(34) &
            "and tipo_documento = " & Chr(34) & lblTipoDoc.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            lblNombre.Text = read.Item(0).ToString
            lblApellido.Text = read.Item(1).ToString
        Catch
            MsgBox("N°Error 308: Error al cargar nombre y apellido del cliente.")
        End Try
        cx.Close()

        'obtener precio categoria-------------------------------------------------------------------------------
        conexion()
        Try
            cm.CommandText = "SELECT nombre_categoria FROM vehiculo where padron = " & Chr(34) & lblPadron.Text & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Categoria = read.Item(0).ToString
        Catch
            MsgBox("N°Error 309: Error al cargar nombre de la categoria.")
        End Try
        cx.Close()

        Dim precioC As Integer
        conexion()
        Try
            cm.CommandText = "SELECT precio FROM categoria where nombre_categoria = " & Chr(34) & Categoria & Chr(34)
            read = cm.ExecuteReader
            read.Read()
            Label13.Text = read.Item(0).ToString * diferencia 'calculo de precio
            LabelDepo.Text = read.Item(0).ToString * diferencia * 0.25 'calculo de deposito 25% del precio
            precioC = read.Item(0)
        Catch
            MsgBox("N°Error 310: Error al cargar precio de la categoria.")
        End Try
        cx.Close()

        'validar si se paso de la fecha de devolucion, en tal caso cobrarle por dia 50%
        'parsear las fechas de las label a datetime para los calculos
        Dim fechaFinal, fechaDevuelto As Date
        Dim dif As Integer 'diferencia de dias a cobrarle

        Dim a As String
        a = lblFFin.Text

        fechaFinal = Date.ParseExact(lblFFin.Text, "dd/MM/yyyy",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

        fechaDevuelto = Date.ParseExact(Label18.Text, "yyyy-MM-dd",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)

        If fechaDevuelto > fechaFinal Then
            dif = DateDiff("d", fechaFinal, fechaDevuelto)
            MsgBox("Por entregar tarde el vehiculo debe pagar una multa de: " & dif * precioC * 1.5)
        End If
    End Sub














    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'Ingresar datos de reserva

        If estado = "Reserva" Then
            Dim alquilado As Boolean = True
            Dim doc2, tipoDoc2 As String
            If lblDocSec.Text = "" And lblTDocSec.Text = "" Then
                doc2 = ""
                tipoDoc2 = ""
            Else
                doc2 = lblDocSec.Text
                tipoDoc2 = lblTDocSec.Text
            End If

            'alquilar
            conexion() 'ingresar datos a reserva
            Try
                cm.CommandText = "insert into reserva values (" & Label3.Text & "," & LabelDepo.Text & "," & Chr(34) & lblDoc.Text & Chr(34) & "," &
                    Chr(34) & lblTipoDoc.Text & Chr(34) & "," & Chr(34) & lblPadron.Text & Chr(34) & "," & Chr(34) & lblFInicio.Text & Chr(34) &
                    "," & Chr(34) & lblFFin.Text & Chr(34) & ",null,null,'Reservado'," & Chr(34) & doc2 & Chr(34) & "," &
                    Chr(34) & tipoDoc2 & Chr(34) & ");"
                cm.ExecuteNonQuery()
                MsgBox("Reservado Correctamente")
            Catch
                alquilado = False
                MsgBox("N°Error 311: Error al reservar.")
            End Try
            cx.Close()
            historial(cm.CommandText)


            'cambiar disponibilidad del vehiculo a Alquilado
            If alquilado = True Then
                conexion()
                Try
                    cm.CommandText = "UPDATE vehiculo SET Disponibilidad = " & Chr(34) & "Reservado" & Chr(34) &
                        " WHERE padron= " & Chr(34) & lblPadron.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch
                    alquilado = False
                    MsgBox("N°Error 312: Error al actualizar disponibilidad del vehiculo.")
                End Try
                cx.Close()
                historial(cm.CommandText)
            End If

            'modificar ubicacion del vehiculo (tabla esta), el traslado-----------------

            If alquilado = True Then
                Dim ids As Integer
                conexion()
                Try
                    cm.CommandText = "SELECT id_sucursal FROM sucursal where nombre_ciudad = " & Chr(34) & lblCiudad.Text & Chr(34)
                    read = cm.ExecuteReader
                    read.Read()
                    ids = read.Item(0)
                Catch

                End Try
                cx.Close()
                historial(cm.CommandText)

                conexion()
                Try
                    cm.CommandText = "UPDATE esta SET id_sucursal = " & ids &
                        " WHERE padron= " & Chr(34) & lblPadron.Text & Chr(34)
                    cm.ExecuteNonQuery()
                Catch

                End Try
                cx.Close()
                historial(cm.CommandText)
            End If

            '---------------------------------------------------------------------------


            'cambiar estado del cliente a Activo
            If alquilado = True Then
                    conexion()
                    Try
                        cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Activo" & Chr(34) &
                    " WHERE documento_cliente = " & Chr(34) & lblDoc.Text & Chr(34) &
                    "and tipo_documento = " & Chr(34) & lblTipoDoc.Text & Chr(34)
                        cm.ExecuteNonQuery()
                    Catch
                    MsgBox("N°Error 313: Error al actualizar estado del Cliente.")
                End Try
                cx.Close()
                historial(cm.CommandText)
            End If

            End If

            If estado = "Alquiler" Then
            'ingresar nueva fecha, la de retiro mediante un modify
            conexion()
            Try
                cm.CommandText = "UPDATE reserva SET fecha_retiro = " & Chr(34) & Label14.Text & Chr(34) &
                " WHERE id_reserva= " & Label3.Text
                cm.ExecuteNonQuery()
            Catch
                MsgBox("N°Error 314: Error al actualizar la fecha de retiro.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            conexion()
            Try
                cm.CommandText = "UPDATE reserva SET estado_reserva = " & Chr(34) & "Alquilado" & Chr(34) &
                " WHERE id_reserva= " & Label3.Text
                cm.ExecuteNonQuery()
                MsgBox("Alquilado Correctamente.")
            Catch
                MsgBox("N°Error 315: Error al actualizar estado de reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            'disponibilidad del vehiculo*-*
            conexion()
            Try
                cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Alquilado" & Chr(34) &
                " WHERE padron= " & Chr(34) & lblPadron.Text & Chr(34)
                cm.ExecuteNonQuery()
            Catch
                MsgBox("ERROR con la operacion")
            End Try
            cx.Close()
        End If

        If estado = "Fin Alquiler" Then
            'ingresar nueva fecha, la de retiro mediante un modify
            conexion()
            Try
                cm.CommandText = "UPDATE reserva SET fecha_devolucion = " & Chr(34) & Label18.Text & Chr(34) &
                " WHERE id_reserva= " & Label3.Text
                cm.ExecuteNonQuery()
            Catch
                MsgBox("N°Error 316: Error al actualizar fecha de devolucion.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            'actualizar estado del alquiler
            conexion()
            Try
                cm.CommandText = "UPDATE reserva SET estado_reserva = " & Chr(34) & "Finalizado" & Chr(34) &
                " WHERE id_reserva= " & Label3.Text
                cm.ExecuteNonQuery()
                MsgBox("Alquiler finalizado Correctamente.")
            Catch
                MsgBox("N°Error 317: Error al actualizar estado de reserva.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            'actualizar datos del cliente
            conexion()
            Try
                cm.CommandText = "UPDATE cliente SET estado = " & Chr(34) & "Inactivo" & Chr(34) &
                " WHERE documento_cliente = " & Chr(34) & lblDoc.Text & Chr(34) & "and tipo_documento = " & Chr(34) & lblTipoDoc.Text & Chr(34)
                cm.ExecuteNonQuery()
            Catch
                MsgBox("N°Error 318: Error al actualizar estado de cliente.")
            End Try
            cx.Close()
            historial(cm.CommandText)

            'actualizar datos del vehiculo
            conexion()
            Try
                cm.CommandText = "UPDATE vehiculo SET disponibilidad = " & Chr(34) & "Disponible" & Chr(34) &
                " WHERE padron = " & Chr(34) & lblPadron.Text & Chr(34)
                cm.ExecuteNonQuery()
            Catch
                MsgBox("N°Error 319: Error al actualizar disponibilidad de vehiculo.")
            End Try
            cx.Close()
            historial(cm.CommandText)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 'cerrar boleta 
        Me.Close()
    End Sub
End Class