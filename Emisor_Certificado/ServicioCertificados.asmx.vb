Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Windows.Forms

Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime


<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class ServicioCertificados
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function GetSQMAdress() As DataSet
        Logger.Write("Recibe llamado GetSQMAdress")

        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("selDireccion", Type.GetType("System.String"))
        dt.Columns.Add("descripcion", Type.GetType("System.String"))
        Dim dr As DataRow = dt.NewRow
        dr("selDireccion") = "CL"
        dr("descripcion") = "SQM SA"
        dt.Rows.Add(dr)
        Dim dr2 As DataRow = dt.NewRow
        dr2("selDireccion") = "EUROPA"
        dr2("descripcion") = "SQM EUROPE NV"
        dt.Rows.Add(dr2)
        Dim dr3 As DataRow = dt.NewRow
        dr3("selDireccion") = "USA"
        dr3("descripcion") = "SQM NORTH AMERICA"
        dt.Rows.Add(dr3)

        ds.Tables.Add(dt)
        Return ds
    End Function

    <WebMethod()> _
    Public Function GetComercialOffices() As DataSet
        Logger.Write("Recibe llamado GetComercialOffices")

        Dim ds As DataSet
        ds = Persistencia.GetComercialOffices

        Logger.Write(" obtuvo datos GetComercialOffices desde BD")
        Return ds
    End Function

    <WebMethod()> _
    Public Function GetCostumerNames() As DataSet
        Logger.Write("Recibe llamado GetCostumerNames")

        Dim ds As DataSet
        ds = Persistencia.GetCostumerNames

        Logger.Write(" obtuvo datos GetCostumerNames desde BD")
        Return ds
    End Function

    <WebMethod()> _
    Public Function GetCertificateSignatures() As DataSet
        Logger.Write("Recibe llamado GetCertificateSignatures")

        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("selLinea", Type.GetType("System.String"))
        dt.Columns.Add("descripcion", Type.GetType("System.String"))
        Dim dr As DataRow = dt.NewRow
        dr("selLinea") = "F"
        dr("descripcion") = "Fertilizer"
        dt.Rows.Add(dr)
        Dim dr2 As DataRow = dt.NewRow
        dr2("selLinea") = "I"
        dr2("descripcion") = "Industrial"
        dt.Rows.Add(dr2)
        ds.Tables.Add(dt)
        Return ds
    End Function

    <WebMethod()> _
   Public Function GetCertificateTo() As DataSet
        Logger.Write("Recibe llamado GetCertificateTo")

        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("id_destino", Type.GetType("System.String"))
        dt.Columns.Add("descripcion", Type.GetType("System.String"))
        Dim dr As DataRow = dt.NewRow
        dr("id_destino") = "0"
        dr("descripcion") = "Costumer"
        dt.Rows.Add(dr)
        Dim dr2 As DataRow = dt.NewRow
        dr2("id_destino") = "1"
        dr2("descripcion") = "Customs"
        dt.Rows.Add(dr2)
        ds.Tables.Add(dt)
        Return ds
    End Function



    <WebMethod()> _
   Public Function SolicitarCertificados(ByVal tipo_destino As Integer, _
                                        ByVal product_comercial_name As String, _
                                        ByVal costumer_name As String, _
                                        ByVal costumer_adress As String, _
                                        ByVal sqm_adress As String, _
                                        ByVal certificate_signature As String, _
                                        ByVal requiere_local_signature As Boolean, _
                                        ByVal oficina_local_signature As String, _
                                        ByVal ver_product_date As Boolean, _
                                        ByVal ver_date_issue As Boolean, _
                                        ByVal ver_certificate_observations As Boolean, _
                                        ByVal certificate_observations As String, _
                                        ByVal envase_inicio As String, _
                                        ByVal envase_fin As String, _
                                        ByVal nro_certificado As Integer, _
                                        ByVal email As String, _
                                        ByVal ver_maxi As Boolean) As String



        Logger.Write("++++++")
        Logger.Write("Recibe llamado SolicitarCertificados")

        Dim result As String = "" 'oficina_local_signature & "-"
        Dim PROCESO_OK As Boolean = True

        '-----VALIDACIONES

        'destinatario
        If tipo_destino <> 0 And tipo_destino <> 1 Then
            Logger.Write("1. destinatario validado = NO")

            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_destinatario")
            result += msg_error & " " '"El campo destinatario debe ser 0 ó 1"
            PROCESO_OK = False
        Else
            Logger.Write("1. destinatario validado = SI")
        End If


        'costumer name
        If Not Persistencia.CostumerNameValidar(costumer_name) Then
            Logger.Write("2. costumer_name validado = NO")

            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_costumer")
            result += msg_error & " " '"El campo Costumer Name ingresado no corresponde"
            PROCESO_OK = False
        Else
            Logger.Write("2. costumer_name validado = SI")
        End If


        'sqm adress
        If Not sqm_adress.Equals("CL") And Not sqm_adress.Equals("EUROPA") And Not sqm_adress.Equals("USA") Then
            Logger.Write("3. sqm_adress validado = NO")

            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_sqm_adress")
            result += msg_error & " " '"El campo SQM ADDRESS ingresado no corresponde"
            PROCESO_OK = False
        Else
            Logger.Write("3. sqm_adress validado = SI")
        End If

        'certificate_signature
        If Not certificate_signature.Equals("F") And Not certificate_signature.Equals("I") Then
            Logger.Write("4. certificate_signature validado = NO")

            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_certif_signat")
            result += msg_error & " " '"El campo Certificate Signature ingresado no corresponde"
            PROCESO_OK = False
        Else
            Logger.Write("4. certificate_signature validado = SI")
        End If

        If requiere_local_signature Then
            Logger.Write("5. requiere_local_signature")
            'oficina_local_siganture  
            If Not (Persistencia.OficinaLocalSignatureValidar(oficina_local_signature)) Then
                Logger.Write("5. requiere_local_signature validado = NO")

                Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_oficina_local_signat")
                result += msg_error & " " '"Se ha indicado el despliegue de Local Signature, pero el nombre de la Oficina Comercial es incorrecto: " & oficina_local_signature.ToString
                PROCESO_OK = False
            Else
                Logger.Write("5. requiere_local_signature validado = SI")
            End If
        End If

        'email valido
        If Not Validar_Email(email) Then
            Logger.Write("6. mail validado = NO")

            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_email")
            result += msg_error & " " '"Formato de correo electronico invalido "
            PROCESO_OK = False
        Else
            Logger.Write("6. mail validado = SI")
        End If

        '-----FIN VALIDACIONES

        Dim ruta_certificado As String = System.Configuration.ConfigurationSettings.AppSettings.Get("ruta_certificado") 'DEL WEBCONFIG 
        Dim ruta_archivos As String = System.Configuration.ConfigurationSettings.AppSettings.Get("ruta_archivos") 'DEL WEBCONFIG

        'en base a los parametros ve cuantos certificados tiene q generar (nro certif o inicio-final)
        Dim certificados As New ArrayList
        If nro_certificado > 0 Then
            Logger.Write("Intenta generar x Certificado")

            'valida que exista el certificado
            If Not Persistencia.ValidaCertificado(nro_certificado) Then
                Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_nro_certif")
                result += msg_error & " " '"No existe el certificado solicitado"
                PROCESO_OK = False
                Logger.Write("Certificado " & nro_certificado & " NO encontrado")
            Else
                Logger.Write("Certificado " & nro_certificado & " encontrado")
            End If
            certificados.Add(nro_certificado)
        Else
            Logger.Write("Intenta generar x Rangos")

            certificados = Persistencia.ObtenerListaCertificados(envase_inicio, envase_fin)
        End If

        If certificados.Count = 0 And PROCESO_OK Then
            Dim msg_error As String = System.Configuration.ConfigurationSettings.AppSettings.Get("msg_cant_certif")
            result += msg_error & " " '"No existen certificados para los parametros ingresados"
            PROCESO_OK = False

            Logger.Write("NO encontro certificados en base al rango")
        Else
            Logger.Write("encontro certificados en base al rango")
        End If

        'por cada certificado agrega una hoja al pdf
        '  Dim theDoc As New WebSupergoo.ABCpdf4.Doc
        Dim thedoc As New ABCpdf4.Doc

        Dim ruta_completa As String = ""
        If PROCESO_OK Then
            For Each certif As Integer In certificados

                Logger.Write("Busqueda info certificado " & certif)

                Try
                    Persistencia.ActualizarConfiguracionCertificado(certif, tipo_destino, product_comercial_name, costumer_name, _
                                                        costumer_adress, sqm_adress, certificate_signature, requiere_local_signature, _
                                                        oficina_local_signature, ver_product_date, ver_date_issue, _
                                                        ver_certificate_observations, certificate_observations, email, ver_maxi)

                    '  theDoc.Page = theDoc.AddPage()
                    ruta_completa = DefineURLCompletaCertificado(ruta_certificado, certif, tipo_destino, product_comercial_name, _
                                                                    costumer_name, costumer_adress, certificate_signature, sqm_adress, _
                                                                    ver_certificate_observations, requiere_local_signature, _
                                                                    oficina_local_signature, ver_date_issue, ver_maxi)

                    Logger.Write("URL que va a llamar: " & ruta_completa)

                    Dim id As Integer = theDoc.AddImageUrl(ruta_completa, True, 0, False)
                    theDoc.Page = theDoc.AddPage()
                    id = theDoc.AddImageToChain(id)

                    'Dim id2 As Integer = theDoc.AddImageUrl(ruta_certificado, True, 0, False)
                    'theDoc.Page = theDoc.AddPage()
                    'id2 = theDoc.AddImageToChain(id2)

                Catch ex As Exception

                    Logger.Write("ERROR en busqueda info certificado " & certif & ": " & ex.Message())

                    result += (" - ERROR CERTIFICADO: " & ex.Message() & " @@ " & ruta_completa & " | ")
                    PROCESO_OK = False
                    Exit For
                End Try
            Next
        End If

        Dim nombre_pdf As String
        If PROCESO_OK Then
            Try
                nombre_pdf = "Certificate_" & certificados(0) & "-" & certificados(certificados.Count - 1)

                Logger.Write("Certificado a generar: " & nombre_pdf)
               
                Dim actual As Integer = 1
                Dim nombre As String = ruta_archivos & nombre_pdf & "_" & actual & ".pdf"
                While System.IO.File.Exists(nombre)
                    actual += 1
                    nombre = ruta_archivos & nombre_pdf & "_" & actual & ".pdf"
                End While

                nombre_pdf = nombre

                Logger.Write("nombre de archivo PDF: " & nombre)

                theDoc.Save(nombre)
                '     theDoc.Dispose()
                theDoc = Nothing

                Logger.Write("PDF generado")

                ' result += " ||| " & ruta_completa
            Catch ex As Exception

                Logger.Write("ERROR PDF: " & ex.Message)

                result += (" - ERROR PDF: " & ex.Message)
                PROCESO_OK = False
            End Try
        End If

        If PROCESO_OK Then
            Try
                Logger.Write("a enviar mail")

                EnviarMail_2(email, nombre_pdf)

                Logger.Write("proceso mail finalizado")
                'EnviarMail_2()
            Catch ex As Exception

                Logger.Write("ERROR EMAIL:" & ex.ToString)

                result += (" - ERROR EMAIL: " & ex.ToString)
            End Try
        End If

        Logger.Write("proceso finalizado")
        Logger.Write("--------")
        Return result

    End Function


    Public Function Validar_Email(ByVal Email As String) As Boolean
        Try
            Dim i As Integer, iLen As Integer, caracter As String
            Dim pos As Integer, bp As Boolean, iPos As Integer, iPos2 As Integer

            Email = Trim$(Email)

            If Email = vbNullString Then
                Exit Function
            End If

            Email = LCase$(Email)
            iLen = Len(Email)

            For i = 1 To iLen
                caracter = Mid(Email, i, 1)

                If (Not (caracter Like "[a-z]")) And (Not (caracter Like "[0-9]")) Then

                    If InStr(1, "_-" & "." & "@", caracter) > 0 Then
                        If bp = True Then
                            Return False
                        Else
                            bp = True

                            If i = 1 Or i = iLen Then
                                Return False
                            End If

                            If caracter = "@" Then
                                If iPos = 0 Then
                                    iPos = i
                                Else
                                    Return False
                                End If
                            End If
                            If caracter = "." Then
                                iPos2 = i
                            End If

                        End If
                    Else
                        Return False
                    End If
                Else
                    bp = False
                End If
            Next i
            If iPos = 0 Or iPos2 = 0 Then
                Return False
            End If

            If iPos2 < iPos Then
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function EnviarMail(ByVal destino As String, ByVal file As String) As Boolean

        Dim host As String = System.Configuration.ConfigurationSettings.AppSettings.Get("servidor_correos")
        Dim puerto_servidor As String = System.Configuration.ConfigurationSettings.AppSettings.Get("puerto_servidor")
        Dim ssl As Integer = System.Configuration.ConfigurationSettings.AppSettings.Get("ssl")
        Dim from As String = System.Configuration.ConfigurationSettings.AppSettings.Get("cuenta_origen")
        Dim pass As String = System.Configuration.ConfigurationSettings.AppSettings.Get("pass")
        Dim asunto As String = System.Configuration.ConfigurationSettings.AppSettings.Get("asunto")
        Dim mensaje As String = System.Configuration.ConfigurationSettings.AppSettings.Get("mensaje")

        Dim MailMsg As New System.Net.Mail.MailMessage(New System.Net.Mail.MailAddress(from.Trim()), New System.Net.Mail.MailAddress(destino))
        MailMsg.BodyEncoding = Encoding.Default
        MailMsg.Subject = asunto.Trim()
        MailMsg.Body = mensaje.Trim() & vbCrLf
        MailMsg.IsBodyHtml = True
        Dim MsgAttach As System.Net.Mail.Attachment
        If Not file = "" Then
            MsgAttach = New System.Net.Mail.Attachment(file)
            MailMsg.Attachments.Add(MsgAttach)
        End If

        Dim SmtpMail As New System.Net.Mail.SmtpClient
        SmtpMail.Host = host
        SmtpMail.Port = puerto_servidor

        SmtpMail.UseDefaultCredentials = False

        SmtpMail.Credentials = New System.Net.NetworkCredential(from, pass)

        If ssl = 1 Then SmtpMail.EnableSsl = True Else SmtpMail.EnableSsl = False

        SmtpMail.Send(MailMsg)
        Try
            MsgAttach.Dispose()
        Catch : End Try
        SmtpMail = Nothing

    End Function

    Public Function EnviarMail_2(ByVal destino As String, ByVal file As String) As Boolean
        Dim servidor As String = System.Configuration.ConfigurationSettings.AppSettings.Get("servidor_correos")
        Dim puerto_servidor As String = System.Configuration.ConfigurationSettings.AppSettings.Get("puerto_servidor")
        Dim ssl As Integer = System.Configuration.ConfigurationSettings.AppSettings.Get("ssl")
        Dim from As String = System.Configuration.ConfigurationSettings.AppSettings.Get("cuenta_origen")
        Dim pass As String = System.Configuration.ConfigurationSettings.AppSettings.Get("pass")
        Dim asunto As String = System.Configuration.ConfigurationSettings.AppSettings.Get("asunto")
        Dim mensaje As String = System.Configuration.ConfigurationSettings.AppSettings.Get("mensaje")

        Dim loConfig
        Dim lcSchema = "http://schemas.microsoft.com/cdo/configuration/"
        loConfig = CreateObject("CDO.Configuration")
        With loConfig.FIELDS
            .ITEM(lcSchema + "smtpserver") = servidor
            .ITEM(lcSchema + "smtpserverport") = puerto_servidor  '465 '&& ó 587
            .ITEM(lcSchema + "sendusing") = 2
            .ITEM(lcSchema + "smtpauthenticate") = True
            .ITEM(lcSchema + "smtpusessl") = IIf(ssl = 1, True, False)
            .ITEM(lcSchema + "sendusername") = from
            .ITEM(lcSchema + "sendpassword") = pass
            .UPDATE()
        End With
        Dim loMsg = CreateObject("CDO.Message")
        With loMsg
            .Configuration = loConfig
            .FROM = from
            .TO = destino
            .Subject = asunto
            .TextBody = mensaje
            .AddAttachment(file)
            Logger.Write("intentando enviar mail")
            .Send()
        End With
    End Function

   

   

    Private Function DefineURLCompletaCertificado(ByVal ruta_completa As String, ByVal ID_Certificado As String, ByVal para As Integer, _
                                    ByVal nombre_comercial As String, ByVal nombre_cliente As String, ByVal direccion_cliente As String, _
                                    ByVal selLinea As String, ByVal selDireccion As String, ByVal mostrar_comentario As String, _
                                    ByVal firma_especial As String, ByVal selFilial As String, ByVal ver_fecha As String, ByVal ver_maxi As String) As String

        Dim rdpara As String = "false"
        If para = 0 Then
            rdpara = "true"
        End If

        Dim rand As New Random
        Dim url As String
        url = "number=" & ID_Certificado
        url = url & "&rdPARA=" & rdpara 'rdPARA
        url = url & "&txtNOMBRE_COMERCIAL=" & nombre_comercial 'txtNOMBRE_COMERCIAL
        url = url & "&txtCustomerName=" & nombre_cliente 'txtCustomerName
        url = url & "&txtCustomerAddress=" & direccion_cliente 'txtCustomerAddress
        url = url & "&selLinea=" & selLinea
        url = url & "&selDireccion=" & selDireccion
        url = url & "&chkMostrarComentario=" & mostrar_comentario 'chkMostrarComentario
        url = url & "&chkFirmaEspecial=" & firma_especial 'chkFirmaEspecial
        url = url & "&selFilial=" & selFilial
        url = url & "&chkVerFecha=" & ver_fecha 'chkVerFecha
        url = url & "&chkVerMaxi=" & ver_maxi 'chkVerFecha
        url = url & "&varClear=" & rand.Next(1, 99999) 'chkVerFecha

        ruta_completa = ruta_completa & "?" & url
        Return ruta_completa
    End Function
End Class

