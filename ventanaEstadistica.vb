Imports System.Data.Odbc
Public Class ventanaEstadistica
    Dim cx As New Odbc.OdbcConnection
    Dim cm As New Odbc.OdbcCommand
    Dim da As New Odbc.OdbcDataAdapter
    Dim ds As New Data.DataSet

    Public Sub conexion() 'prueba la conexion a la base de datos
        Try
            cx.ConnectionString = "FileDsn=" & Application.StartupPath & "\odbcCharm.dsn;UID=" & LoginSU.user & ";PWD=" & LoginSU.contraseña
            cx.Open()
            cm.Connection = cx
        Catch
            MsgBox("N°Error 320: No hay conexion con la base de datos.")
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
            MsgBox("N°Error 321: Error al mantener historial.")
        End Try
        cx.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = Not 0 Then
            DataGridView1.DataSource.clear()
        End If
        conexion()
        Try
            cm.CommandText = "select count(reserva.id_reserva) cantidad, month(reserva.fecha_inicio) mes from reserva group by mes"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridView1.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 322: Al cargar consulta.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount = Not 0 Then
            DataGridView1.DataSource.clear()
        End If
        conexion()
        Try
            cm.CommandText = "select * from historial"
            cm.ExecuteNonQuery()
            da.SelectCommand = cm
            da.Fill(ds, "Vehiculos")
            DataGridView1.DataSource = ds.Tables("Vehiculos")
        Catch
            MsgBox("N°Error 323: Al cargar historial.")
        End Try
        cx.Close()
        historial(cm.CommandText)
    End Sub
End Class