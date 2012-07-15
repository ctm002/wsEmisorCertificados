Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OracleClient

'Imports mshtml 'No olvidar esta referencia

Public Class Persistencia

    Public Shared Function GetComercialOffices() As DataSet
        Try
            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim ds As New DataSet

            Dim sql As String = "SELECT MFI_COD_FILIAL CODIGO, initcap(MFI_DESCRIP_L) DESCRIPCION " & _
                            " FROM SGC_MAE_FILIALES WHERE MFI_TIPO = 'E' OR MFI_COD_FILIAL = 1 ORDER BY MFI_DESCRIP_L"

            cmd.Connection = conn
            cmd.CommandText = sql

            Dim da As New OracleDataAdapter(cmd)

            da.Fill(ds, "Consulta")
            conn.Close()
            Return ds
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function GetCostumerNames() As DataSet
        Try
            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim ds As New DataSet

            Dim sql As String = "select CODIGO, RAZON_SOCIAL from SGC_TB_CLIENTES"

            cmd.Connection = conn
            cmd.CommandText = sql
            Dim da As New OracleDataAdapter(cmd)
            da.Fill(ds, "Consulta")
            conn.Close()
            Return ds
        Catch ex As Exception
            Throw ex
        End Try
    End Function


   
    Public Shared Function ValidaCertificado(ByVal nro_certificado As Integer) As Boolean
        Try

            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim adaptador As New OracleDataAdapter
            Dim ds As New DataSet

            Dim sql As String = " select ID_CERTIFICADO from CER_TB_CERTIFICADOS WHERE ID_CERTIFICADO = " & nro_certificado

            cmd.Connection = conn
            cmd.CommandText = sql
            Dim da As New OracleDataAdapter(cmd)
            da.Fill(ds, "Consulta")
            conn.Close()
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function ObtenerListaCertificados(ByVal envase_inicio As Integer, ByVal envase_fin As Integer) As ArrayList
        Try
            Dim certificados As New ArrayList()

            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim adaptador As New OracleDataAdapter
            Dim ds As New DataSet

            Dim sql As String = " select ID_CERTIFICADO from CER_TB_RANGO_ENVASES " & _
                                " where 	( INICIO <= " & envase_inicio & " and FIN >= " & envase_inicio & " ) or  	/*primer certifi*/ " & _
                                "	( INICIO <= " & envase_fin & " and FIN >= " & envase_fin & " ) or		/*ultimo certif*/ " & _
                                "	( INICIO > " & envase_inicio & " and FIN < " & envase_fin & " ) or		/*entre medio*/ " & _
                                "	( INICIO <= " & envase_inicio & " and FIN >= " & envase_fin & " )		/*rango menor*/"

            cmd.Connection = conn
            cmd.CommandText = sql
            Dim da As New OracleDataAdapter(cmd)
            da.Fill(ds, "Consulta")
            For Each dr As DataRow In ds.Tables(0).Rows
                certificados.Add(dr(0))
            Next
            conn.Close()
            Return certificados
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Shared Function ActualizarConfiguracionCertificado(ByVal nro_certif, ByVal tipo_destino, ByVal product_comercial_name, ByVal costumer_name, _
                                                    ByVal costumer_adress, ByVal sqm_adress, ByVal certificate_signature, ByVal requiere_special_local_signature, _
                                                    ByVal oficina_local_siganture, ByVal ver_product_date, ByVal ver_date_issue, _
                                                    ByVal ver_certificate_observations, ByVal certificate_observations, ByVal email, ByVal ver_maxi) As Boolean
        Dim conn As OracleConnection
        Dim transaction As OracleTransaction
        Dim sql_insert As String = ""
        Try
            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            conn = New OracleConnection(sqldb)
            conn.Open()

            transaction = conn.BeginTransaction


            'actualiza el certificado

            Dim sql_update As String = "UPDATE CER_TB_CERTIFICADOS SET OBSERVACIONES = '" & certificate_observations & "' WHERE ID_CERTIFICADO = " & nro_certif

            Dim cmd As New OracleCommand
            cmd.CommandType = System.Data.CommandType.Text
            cmd.CommandText = sql_update
            cmd.Connection = conn

            cmd.Transaction = transaction
            cmd.ExecuteNonQuery()


            'obtiene el numero de versiones de config del certif

            Dim ds As New DataSet
            Dim sql As String = "SELECT nvl(max(version), 0) + 1 Version FROM CER_TB_CONFIG_USER WHERE CERTIFICATE = " & nro_certif
            Dim cmd2 As New OracleCommand
            cmd2.Connection = conn
            cmd2.CommandText = sql
            cmd2.Transaction = transaction
            Dim da As New OracleDataAdapter(cmd2)
            da.Fill(ds)
            Dim version As Integer
            If ds.Tables(0).Rows.Count > 0 Then
                version = ds.Tables(0).Rows(0)("Version")
            End If


            ' inserta la nueva configuracion del certif 

            sql_insert = "insert into cer_tb_config_user "
            sql_insert = sql_insert & "(to_destination, "
            sql_insert = sql_insert & " Commercial_name, "
            sql_insert = sql_insert & " Customer_name, "
            sql_insert = sql_insert & " Customer_Address,"
            sql_insert = sql_insert & " sqm_Address, "
            sql_insert = sql_insert & " special_signature,"
            sql_insert = sql_insert & " signature_value, "
            sql_insert = sql_insert & " view_date,"
            sql_insert = sql_insert & " view_comments,"
            sql_insert = sql_insert & " certificate,"
            sql_insert = sql_insert & " name_user, "
            sql_insert = sql_insert & " version, "
            sql_insert = sql_insert & " comments, "
            If requiere_special_local_signature Then
                sql_insert = sql_insert & " filial, "
            End If
            sql_insert = sql_insert & " VIEW_PRODUCTION_DATE, "
            sql_insert = sql_insert & " VIEW_MAXI, "
            sql_insert = sql_insert & " email )"
            sql_insert = sql_insert & " VALUES "
            sql_insert = sql_insert & "(" & tipo_destino & ", "
            sql_insert = sql_insert & " '" & product_comercial_name & "', "
            sql_insert = sql_insert & " '" & costumer_name & "', "
            sql_insert = sql_insert & " '" & costumer_adress & "',"
            sql_insert = sql_insert & " '" & sqm_adress & "', "
            sql_insert = sql_insert & " " & IIf(requiere_special_local_signature, 1, 0) & ","
            sql_insert = sql_insert & " '" & certificate_signature & "',"
            sql_insert = sql_insert & " " & IIf(ver_date_issue, 1, 0) & ","
            sql_insert = sql_insert & " " & IIf(ver_certificate_observations, 1, 0) & ","
            sql_insert = sql_insert & " " & nro_certif & ","
            sql_insert = sql_insert & " '" & " " & "', "
            sql_insert = sql_insert & " " & version & ","
            sql_insert = sql_insert & " '" & certificate_observations & "',"
            If requiere_special_local_signature Then
                sql_insert = sql_insert & " " & oficina_local_siganture & ","
            End If
            sql_insert = sql_insert & " " & IIf(ver_product_date, 1, 0) & ","
            sql_insert = sql_insert & " " & IIf(ver_maxi, 1, 0) & ","
            sql_insert = sql_insert & " '" & email & "' )"


            Dim cmd_insert As New OracleCommand
            cmd_insert.CommandType = System.Data.CommandType.Text
            cmd_insert.CommandText = sql_insert
            cmd_insert.Connection = conn

            cmd_insert.Transaction = transaction
            cmd_insert.ExecuteNonQuery()

            transaction.Commit()
            conn.Close()

            Return True

        Catch ex As Exception
            Try
                transaction.Rollback()
                conn.Close()
            Catch ex2 As Exception

            End Try
            Throw New Exception(ex.Message & " -- " & sql_insert)
        End Try
    End Function

    Public Shared Function CostumerNameValidar(ByVal costumer_name As String) As Boolean
        Try
            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim ds As New DataSet

            Dim sql As String = " select CODIGO  from SGC_TB_CLIENTES" & _
                    " where RAZON_SOCIAL = '" & costumer_name & "'"

            cmd.Connection = conn
            cmd.CommandText = sql
            Dim da As New OracleDataAdapter(cmd)
            da.Fill(ds, "Consulta")
            conn.Close()
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function OficinaLocalSignatureValidar(ByVal oficina_local_signature As String) As Boolean
        Try
            Dim sqldb As String = System.Configuration.ConfigurationSettings.AppSettings.Get("string_conexion")

            Dim conn As New OracleConnection(sqldb)
            conn.Open()
            Dim cmd As New OracleCommand
            Dim ds As New DataSet

            Dim sql As String = "SELECT MFI_COD_FILIAL FROM SGC_MAE_FILIALES WHERE MFI_COD_FILIAL = " & oficina_local_signature

            cmd.Connection = conn
            cmd.CommandText = sql
            Dim da As New OracleDataAdapter(cmd)
            da.Fill(ds, "Consulta")
            conn.Close()
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function


End Class


