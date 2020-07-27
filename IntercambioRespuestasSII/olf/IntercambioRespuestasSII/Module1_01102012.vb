Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Collections.Generic
Imports System.Net
Imports System.Web.Services.Protocols
Imports RespuestasSII






Module Module1
    Public TpoDoc As String
    Public FchEmi As String
    Public RutEmi As String
    Public TpoDte As String
    Public Estado As String
    Public Llave As String
    Public oXml As XmlDocument
    Public xObj As XmlElement
    Public sObj As String
    Public arrArch() As String
    Public sqlCnx As New SqlConnection
    Public sqlCmd As New SqlCommand

    Public Cdgintrecep As String
    Public Cdgvendedor As String
    Public Ciudadorigen As String
    Public Ciudadrecep As String
    Public Cmnaorigen As String
    Public Cmnarecep As String
    Public Correoemisor As String
    Public Correorecep As String
    Public Dirorigen As String
    Public Dirrecep As String
    Public Fchemis As String
    Public Fchvenc As String
    Public Fmapago As String
    Public folio As String
    Public Giroemis As String
    Public Girorecep As String
    Public Imptoreten As String
    Public Iva As String
    Public Ivanoret As String
    Public Mntexe As String
    Public Mntneto As String
    Public Mnttotal As String
    Public Montonf As String
    Public Montoperiodo As String
    Public Referencia As String
    Public Rutemisor As String
    Public Rutrecep As String
    Public Rznsoc As String
    Public Rznsocrecep As String
    Public Saldoanterior As String
    Public Tasaiva As String
    Public Tipodte As String
    Public Tpomoneda As String
    Public Valcomexe As String
    Public Valcomiva As String
    Public Valcomneto As String
    Public Vlrpagar As String
    Sub Main()

        '
        '
        '
        Dim CertPath As String = My.Settings.CertPath
        Dim CertPass As String = My.Settings.CertPass
        Dim RequiereProxy As String = My.Settings.RequiereProxy
        Dim ProxyServer As String = My.Settings.ProxyServer
        Dim ProxyUser As String = My.Settings.ProxyUser
        Dim ProxyPass As String = My.Settings.ProxyPass
        Dim DtePath As String = My.Settings.DtePath
        Dim DtePathProc As String = My.Settings.DtePathProcOK
        Dim maximoIntentos As String = My.Settings.maximoIntentos
        Dim DtePathProcNOK As String = My.Settings.DtePathProcNOK
        Dim VidaSemillaMil As String = My.Settings.VidaSemillaMil
        Dim DteEstOK As String = My.Settings.DteEstOK

        Dim rp As New RespuestasSII.RespuestasSII
        Dim resultado As String()
        'retorna
        'cantidad total
        'cantidad de procesados correctamente
        resultado = rp.Proceso(CertPath, CertPass, RequiereProxy, ProxyServer, ProxyUser, ProxyPass, DtePath, _
                               DtePathProc, maximoIntentos, DtePathProcNOK, VidaSemillaMil, DteEstOK)
        Dim cantDte As Integer, cantProc As Integer

        cantDte = resultado(0)
        cantProc = resultado(1)

        ' string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
        '

        Dim ht As New Hashtable
        Dim rutaIN As String = System.Configuration.ConfigurationSettings.AppSettings.Get("IN")
        'Dim aaa As New StreamWriter("C:\in.txt")
        arrArch = Directory.GetFiles(DtePathProc, "*.xml")
       
        'For Each sObj In arrArch

        '  Console.Clear()
        ' Console.WriteLine("Procesando DTE, " & sObj)
        '    oXml = New XmlDocument()
        '    oXml.Load(sObj)
        '    xObj = oXml.GetElementsByTagName("Documento").Item(0)
        '    If xObj Is Nothing Then
        '        xObj = oXml.GetElementsByTagName("Liquidacion").Item(0)
        '    End If
        '    RutEmi = xObj.GetElementsByTagName("RUTEmisor").Item(0).InnerText
        '    TpoDoc = xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText
        '    folio = xObj.GetElementsByTagName("Folio").Item(0).InnerText

        '    Llave = RutEmi.Split("-")(0) & TpoDoc.PadLeft(3, "0") & folio.PadLeft(10, "0")

        '    If Not ht.ContainsKey(Llave.ToString) Then
        '        ht.Add(Llave, sObj)
        '        'aaa.WriteLine(Llave & "-" & sObj)
        '        'aaa.Flush()
        '    End If
        'Next

        'leer carpeta out

        Dim rutaOUT As String = System.Configuration.ConfigurationSettings.AppSettings.Get("OUT")
        arrArch = Directory.GetFiles(rutaOUT, "*.xml")
        Dim aou As New StreamWriter(Path.GetDirectoryName(rutaOUT) & "\Log_OK.txt", False, System.Text.Encoding.GetEncoding("iso-8859-1"))
        Dim aou1 As New StreamWriter(Path.GetDirectoryName(rutaOUT) & "\Log_Error.txt", False, System.Text.Encoding.GetEncoding("iso-8859-1"))

        For Each sObj In arrArch
            'Dim arr As String = Path.GetFileName(sObj).Split("_")(1)

            'oXml = New XmlDocument()
            'oXml.Load(sObj)

            'xObj = oXml.GetElementsByTagName("SII:RESP_HDR").Item(0)
            'Estado = xObj.GetElementsByTagName("GLOSA_ESTADO").Item(0).InnerText

            'If Estado = "DTE No Recibido" Then

            '    ' '' ''enviar mail a proveedor
            '    '' ''sqlCnx.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings.Get("BD").ToString()
            '    '' ''sqlCnx.Open()
            '    '' ''sqlCmd.Connection = sqlCnx
            '    '' ''sqlCmd.CommandTimeout = 999999999
            '    '' ''sqlCmd.CommandType = System.Data.CommandType.Text
            '    '' ''sqlCmd.CommandText = "select mailintercambio from dbo.MailReceptor where rut LIKE '%" + arr + "%'"
            '    '' ''Dim mail As Object = sqlCmd.ExecuteScalar()
            '    '' ''sqlCnx.Close()


            '    '' ''Dim DTE_Error As String = Path.GetFileNameWithoutExtension(sObj).Split("_")(1) & Path.GetFileNameWithoutExtension(sObj).Split("_")(2).PadLeft(3, "0") & Path.GetFileNameWithoutExtension(sObj).Split("_")(3).PadLeft(10, "0")


            '    '' ''oXml = New XmlDocument()
            '    '' ''oXml.Load(ht(DTE_Error.ToString))
            '    '' ''xObj = oXml.GetElementsByTagName("Documento").Item(0)
            '    '' ''If xObj Is Nothing Then
            '    '' ''    xObj = oXml.GetElementsByTagName("Liquidacion").Item(0)
            '    '' ''End If
            '    '' ''FchEmi = xObj.GetElementsByTagName("FchEmis").Item(0).InnerText
            '    '' ''TpoDoc = xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText
            '    '' ''folio = xObj.GetElementsByTagName("Folio").Item(0).InnerText

            '    '' ''Dim email As New MailMessage()
            '    '' ''Dim Host As String = "mail.jordan.cl"
            '    '' ''Dim objSmtp As New SmtpClient(Host)
            '    '' ''email.From = New MailAddress(System.Configuration.ConfigurationSettings.AppSettings.Get("Emisor"))
            '    '' ''email.To.Add(mail)
            '    '' ''email.Subject = "Factura Electrónica " & folio & " rechazada"
            '    '' ''email.Body = "Estimado Proveedor," & vbCrLf & vbCrLf & "Informamos a Ud. que el siguiente Documento Tributario Electronico corresponde a:" & vbCrLf & vbCrLf & "Folio:" & folio + _
            '    '' ''vbCrLf & vbCrLf & "Tipo Dte:" & TpoDoc & vbCrLf & vbCrLf & "Fecha Emision:" & FchEmi & vbCrLf & vbCrLf & "Atte." & vbCrLf & vbCrLf & "BCI."

            '    '' ''If ht.ContainsKey(DTE_Error.ToString) Then
            '    '' ''    'Dim adjunto As New Attachment(ht(Path.GetFileNameWithoutExtension(DTE_Error)))
            '    '' ''    'email.Attachments.Add(adjunto)
            '    '' ''    'Dim adjunto1 As New Attachment(sObj)
            '    '' ''    'email.Attachments.Add(adjunto1)
            '    '' ''End If
            '    ' '' ''Dim adjunto As New Attachment(sObj)
            '    ' '' ''email.Attachments.Add(adjunto)
            '    '' ''Try
            '    '' ''    objSmtp.EnableSsl = False
            '    '' ''    objSmtp.Credentials = New System.Net.NetworkCredential("intercambiobci@jordan.cl", "dte2007")
            '    '' ''    objSmtp.Port = 25
            '    '' ''    objSmtp.Send(email)
            '    '' ''Catch ex As Exception

            '    '' ''End Try
            '    '' ''email.Attachments.Clear()

            'Else

            oXml = New XmlDocument()
            oXml.Load(sObj)

            'If Path.GetFileNameWithoutExtension(sObj).Split("_")(1) & Path.GetFileNameWithoutExtension(sObj).Split("_")(2) & Path.GetFileNameWithoutExtension(sObj).Split("_")(3) = "771225900520000020858" Then
            '    Console.WriteLine("")
            'End If


            xObj = oXml.GetElementsByTagName("Documento").Item(0)
            If xObj Is Nothing Then
                xObj = oXml.GetElementsByTagName("Liquidacion").Item(0)
            End If


            Dim WsBCI_krga As New WebReference.CargaDte
            'Dim WsBCI_Envio As New WebReference.zservicios_fe_v1
            Dim WsBCI_Envio As New WebReference.zservicios_fe_v2

            Dim WsBCI_Detalle As New WebReference.ZfeDetalle
            Dim WsBCI_Impuesto As New WebReference.ZfeImpuestos
            Dim WsBCI_Referencia As New WebReference.ZfeReferencia

            Dim objCredential As New System.Net.NetworkCredential()

            objCredential.UserName = "ws_test"
            objCredential.Password = "banco12"

            WsBCI_Envio.Credentials = objCredential

            'Dim det(xObj.GetElementsByTagName("Detalle").Count - 1) As WebReference.ZfeDetalle
            'det(xObj.GetElementsByTagName("Detalle").Count - 1) = New WebReference.ZfeDetalle

            Dim det(xObj.GetElementsByTagName("Detalle").Count - 1) As WebReference.ZfeDetalle
            Dim ConDet As Integer = 0
            For ConDet = 0 To xObj.GetElementsByTagName("Detalle").Count - 1
                det(ConDet) = New WebReference.ZfeDetalle
                If oXml.OuterXml.Contains("DscItem") = True Then
                    Try
                        det(ConDet).Dscitem = xObj.GetElementsByTagName("DscItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Dscitem = ""
                    End Try

                Else
                    det(ConDet).Dscitem = ""
                End If

                If oXml.OuterXml.Contains("<MontoItem>") = True Then
                    Try
                        det(ConDet).Montoitem = xObj.GetElementsByTagName("MontoItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Montoitem = ""
                    End Try
                Else
                    det(ConDet).Montoitem = ""
                End If

                If oXml.OuterXml.Contains("<NmbItem>") = True Then
                    Try
                        det(ConDet).Nmbitem = xObj.GetElementsByTagName("NmbItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Nmbitem = ""
                    End Try
                Else
                    det(ConDet).Nmbitem = ""
                End If

                If oXml.OuterXml.Contains("<NroLinDet>") = True Then
                    Try
                        det(ConDet).Nrolindet = xObj.GetElementsByTagName("NroLinDet").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Nrolindet = ""
                    End Try
                Else
                    det(ConDet).Nrolindet = ""
                End If

                If oXml.OuterXml.Contains("<PrcItem>") = True Then
                    Try
                        det(ConDet).Prcitem = xObj.GetElementsByTagName("PrcItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Prcitem = ""
                    End Try

                Else
                    det(ConDet).Prcitem = ""
                End If
                If oXml.OuterXml.Contains("<QtyItem>") = True Then
                    Try
                        det(ConDet).Qtyitem = xObj.GetElementsByTagName("QtyItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Qtyitem = ""
                    End Try
                Else
                    det(ConDet).Qtyitem = ""
                End If

                If oXml.OuterXml.Contains("<TpoCodigo>") = True Then
                    Try
                        det(ConDet).Tpocodigo = xObj.GetElementsByTagName("TpoCodigo").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Tpocodigo = ""
                    End Try
                Else
                    det(ConDet).Tpocodigo = ""
                End If

                If oXml.OuterXml.Contains("<UnmdItem>") = True Then
                    Try
                        det(ConDet).Unmditem = xObj.GetElementsByTagName("UnmdItem").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Unmditem = ""
                    End Try
                Else
                    det(ConDet).Unmditem = ""
                End If

                If oXml.OuterXml.Contains("<VlrCodigo>") = True Then
                    Try
                        det(ConDet).Vlrcodigo = xObj.GetElementsByTagName("VlrCodigo").Item(ConDet).InnerText
                    Catch ex As Exception
                        det(ConDet).Vlrcodigo = ""
                    End Try
                Else
                    det(ConDet).Vlrcodigo = ""
                End If

            Next
            WsBCI_krga.Detalle = det


            Dim ConImp As Integer
            Dim Imp(xObj.GetElementsByTagName("ImptoReten").Count - 1) As WebReference.ZfeImpuestos
            For ConImp = 0 To xObj.GetElementsByTagName("ImptoReten").Count - 1
                Imp(ConImp) = New WebReference.ZfeImpuestos

                If oXml.OuterXml.Contains("<MontoImp>") = True Then
                    Try
                        Imp(ConImp).Montoimp = xObj.GetElementsByTagName("MontoImp").Item(ConImp).InnerText
                    Catch ex As Exception
                        Imp(ConImp).Montoimp = ""
                    End Try
                Else
                    Imp(ConImp).Montoimp = ""
                End If

                If oXml.OuterXml.Contains("<TasaImp>") = True Then
                    Try
                        Imp(ConImp).Tasaimp = xObj.GetElementsByTagName("TasaImp").Item(ConImp).InnerText
                    Catch ex As Exception
                        Imp(ConImp).Tasaimp = ""
                    End Try
                Else
                    Imp(ConImp).Tasaimp = ""
                End If

                If oXml.OuterXml.Contains("<TipoImp>") = True Then
                    Try
                        Imp(ConImp).Tipoimp = xObj.GetElementsByTagName("TipoImp").Item(ConImp).InnerText
                    Catch ex As Exception
                        Imp(ConImp).Tipoimp = ""
                    End Try
                Else
                    Imp(ConImp).Tipoimp = ""
                End If

            Next
            WsBCI_krga.Imptoreten = Imp

            Dim ConRef As Integer
            Dim Ref(xObj.GetElementsByTagName("Referencia").Count - 1) As WebReference.ZfeReferencia
            For ConRef = 0 To xObj.GetElementsByTagName("Referencia").Count - 1
                Ref(ConRef) = New WebReference.ZfeReferencia
                If oXml.OuterXml.Contains("<CodRef>") = True Then
                    Try
                        Ref(ConRef).Codref = xObj.GetElementsByTagName("CodRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Codref = ""
                    End Try
                Else
                    Ref(ConRef).Codref = ""
                End If

                If oXml.OuterXml.Contains("<FchRef>") = True Then
                    Try
                        Ref(ConRef).Fchref = xObj.GetElementsByTagName("FchRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Fchref = ""
                    End Try
                Else
                    Ref(ConRef).Fchref = ""
                End If

                If oXml.OuterXml.Contains("<FolioRef>") = True Then
                    Try
                        Ref(ConRef).Folioref = xObj.GetElementsByTagName("FolioRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Folioref = ""
                    End Try
                Else
                    Ref(ConRef).Folioref = ""
                End If

                If oXml.OuterXml.Contains("<IndGlobal>") = True Then
                    Try
                        Ref(ConRef).Indglobal = xObj.GetElementsByTagName("IndGlobal").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Indglobal = ""
                    End Try
                Else
                    Ref(ConRef).Indglobal = ""
                End If
                If oXml.OuterXml.Contains("<NroLinRef>") = True Then
                    Try
                        Ref(ConRef).Nrolinref = xObj.GetElementsByTagName("NroLinRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Nrolinref = ""
                    End Try
                Else
                    Ref(ConRef).Nrolinref = ""
                End If

                If oXml.OuterXml.Contains("<RazonRef>") = True Then
                    Try
                        Ref(ConRef).Razonref = xObj.GetElementsByTagName("RazonRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Razonref = ""
                    End Try
                Else
                    Ref(ConRef).Razonref = ""
                End If
                If oXml.OuterXml.Contains("<TpoDocRef>") = True Then
                    Try
                        Ref(ConRef).Tpodocref = xObj.GetElementsByTagName("TpoDocRef").Item(ConRef).InnerText
                    Catch ex As Exception
                        Ref(ConRef).Tpodocref = ""
                    End Try
                Else
                    Ref(ConRef).Tpodocref = ""
                End If
            Next
            WsBCI_krga.Referencia = Ref

            'Activar WS cargadte
            If oXml.OuterXml.Contains("<CdgIntRecep>") = True Then
                WsBCI_krga.Cdgintrecep = xObj.GetElementsByTagName("CdgIntRecep").Item(0).InnerText
            Else
                WsBCI_krga.Cdgintrecep = ""
            End If

            If oXml.OuterXml.Contains("<Cdgvendedor>") = True Then
                WsBCI_krga.Cdgvendedor = xObj.GetElementsByTagName("Cdgvendedor").Item(0).InnerText
            Else
                WsBCI_krga.Cdgvendedor = ""
            End If
            If oXml.OuterXml.Contains("<CiudadOrigen>") = True Then
                WsBCI_krga.Ciudadorigen = xObj.GetElementsByTagName("CiudadOrigen").Item(0).InnerText
            Else
                WsBCI_krga.Ciudadorigen = ""
            End If

            If oXml.OuterXml.Contains("<CiudadRecep>") = True Then
                WsBCI_krga.Ciudadrecep = xObj.GetElementsByTagName("CiudadRecep").Item(0).InnerText
            Else
                WsBCI_krga.Ciudadrecep = ""
            End If

            If oXml.OuterXml.Contains("<CmnaOrigen>") = True Then
                WsBCI_krga.Cmnaorigen = xObj.GetElementsByTagName("CmnaOrigen").Item(0).InnerText
            Else
                WsBCI_krga.Cmnaorigen = ""
            End If

            If oXml.OuterXml.Contains("<CmnaRecep>") = True Then
                WsBCI_krga.Cmnarecep = xObj.GetElementsByTagName("CmnaRecep").Item(0).InnerText
            Else
                WsBCI_krga.Cmnarecep = ""
            End If

            If oXml.OuterXml.Contains("<CorreoEmisor>") = True Then
                WsBCI_krga.Correoemisor = xObj.GetElementsByTagName("CorreoEmisor").Item(0).InnerText
            Else
                WsBCI_krga.Correoemisor = ""
            End If

            If oXml.OuterXml.Contains("<CorreoRecep>") = True Then
                WsBCI_krga.Correorecep = xObj.GetElementsByTagName("CorreoRecep").Item(0).InnerText
            Else
                WsBCI_krga.Correorecep = ""
            End If

            If oXml.OuterXml.Contains("<DirOrigen>") = True Then
                WsBCI_krga.Dirorigen = xObj.GetElementsByTagName("DirOrigen").Item(0).InnerText
            Else
                WsBCI_krga.Dirorigen = ""
            End If

            If oXml.OuterXml.Contains("<DirRecep>") = True Then
                WsBCI_krga.Dirrecep = xObj.GetElementsByTagName("DirRecep").Item(0).InnerText
            Else
                WsBCI_krga.Dirrecep = ""
            End If

            If oXml.OuterXml.Contains("<FchEmis>") = True Then
                WsBCI_krga.Fchemis = xObj.GetElementsByTagName("FchEmis").Item(0).InnerText
            Else
                WsBCI_krga.Fchemis = ""
            End If

            If oXml.OuterXml.Contains("<FchVenc>") = True Then
                WsBCI_krga.Fchvenc = xObj.GetElementsByTagName("FchVenc").Item(0).InnerText
            Else
                WsBCI_krga.Fchvenc = ""
            End If

            If oXml.OuterXml.Contains("<FmaPago>") = True Then
                WsBCI_krga.Fmapago = xObj.GetElementsByTagName("FmaPago").Item(0).InnerText
            Else
                WsBCI_krga.Fmapago = ""
            End If

            If oXml.OuterXml.Contains("<Folio>") = True Then
                WsBCI_krga.Folio = xObj.GetElementsByTagName("Folio").Item(0).InnerText
            Else
                WsBCI_krga.Folio = ""
            End If

            If oXml.OuterXml.Contains("<GiroEmis>") = True Then
                WsBCI_krga.Giroemis = xObj.GetElementsByTagName("GiroEmis").Item(0).InnerText
            Else
                WsBCI_krga.Giroemis = ""
            End If

            If oXml.OuterXml.Contains("<GiroRecep>") = True Then
                WsBCI_krga.Girorecep = xObj.GetElementsByTagName("GiroRecep").Item(0).InnerText
            Else
                WsBCI_krga.Girorecep = ""
            End If

            If oXml.OuterXml.Contains("<IVA>") = True Then
                WsBCI_krga.Iva = xObj.GetElementsByTagName("IVA").Item(0).InnerText
            Else
                WsBCI_krga.Iva = ""
            End If

            If oXml.OuterXml.Contains("<IVANoRet>") = True Then
                WsBCI_krga.Ivanoret = xObj.GetElementsByTagName("IVANoRet").Item(0).InnerText
            Else
                WsBCI_krga.Ivanoret = ""
            End If

            If oXml.OuterXml.Contains("<MntExe>") = True Then
                WsBCI_krga.Mntexe = xObj.GetElementsByTagName("MntExe").Item(0).InnerText
            Else
                WsBCI_krga.Mntexe = ""
            End If

            If oXml.OuterXml.Contains("<MntNeto>") = True Then
                WsBCI_krga.Mntneto = xObj.GetElementsByTagName("MntNeto").Item(0).InnerText
            Else
                WsBCI_krga.Mntneto = ""
            End If

            If oXml.OuterXml.Contains("<MntTotal>") = True Then
                WsBCI_krga.Mnttotal = xObj.GetElementsByTagName("MntTotal").Item(0).InnerText
            Else
                WsBCI_krga.Mnttotal = ""
            End If

            If oXml.OuterXml.Contains("<MontoNF>") = True Then
                WsBCI_krga.Montonf = xObj.GetElementsByTagName("MontoNF").Item(0).InnerText
            Else
                WsBCI_krga.Montonf = ""
            End If

            If oXml.OuterXml.Contains("<MontoPeriodo>") = True Then
                WsBCI_krga.Montoperiodo = xObj.GetElementsByTagName("MontoPeriodo").Item(0).InnerText
            Else
                WsBCI_krga.Montoperiodo = ""
            End If

            If oXml.OuterXml.Contains("<RUTEmisor>") = True Then
                WsBCI_krga.Rutemisor = xObj.GetElementsByTagName("RUTEmisor").Item(0).InnerText
            Else
                WsBCI_krga.Rutemisor = ""
            End If

            If oXml.OuterXml.Contains("<RUTRecep>") = True Then
                WsBCI_krga.Rutrecep = xObj.GetElementsByTagName("RUTRecep").Item(0).InnerText
            Else
                WsBCI_krga.Rutrecep = ""
            End If

            If oXml.OuterXml.Contains("<RznSoc>") = True Then
                WsBCI_krga.Rznsoc = xObj.GetElementsByTagName("RznSoc").Item(0).InnerText
            Else
                WsBCI_krga.Rznsoc = ""
            End If

            If oXml.OuterXml.Contains("<RznSocRecep>") = True Then
                WsBCI_krga.Rznsocrecep = xObj.GetElementsByTagName("RznSocRecep").Item(0).InnerText
            Else
                WsBCI_krga.Rznsocrecep = ""
            End If
            If oXml.OuterXml.Contains("<SaldoAnterior>") = True Then
                WsBCI_krga.Saldoanterior = xObj.GetElementsByTagName("SaldoAnterior").Item(0).InnerText
            Else
                WsBCI_krga.Saldoanterior = ""
            End If

            If oXml.OuterXml.Contains("<TasaIVA>") = True Then
                WsBCI_krga.Tasaiva = xObj.GetElementsByTagName("TasaIVA").Item(0).InnerText
            Else
                WsBCI_krga.Tasaiva = ""
            End If

            If oXml.OuterXml.Contains("<TipoDTE>") = True Then
                WsBCI_krga.Tipodte = xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText
            Else
                WsBCI_krga.Tipodte = ""
            End If

            If oXml.OuterXml.Contains("<TpoMoneda>") = True Then
                WsBCI_krga.Tpomoneda = xObj.GetElementsByTagName("TpoMoneda").Item(0).InnerText
            Else
                WsBCI_krga.Tpomoneda = ""
            End If

            If oXml.OuterXml.Contains("<ValComExe>") = True Then
                WsBCI_krga.Valcomexe = xObj.GetElementsByTagName("ValComExe").Item(0).InnerText
            Else
                WsBCI_krga.Valcomexe = ""
            End If

            If oXml.OuterXml.Contains("<ValComIVA>") = True Then
                WsBCI_krga.Valcomiva = xObj.GetElementsByTagName("ValComIVA").Item(0).InnerText
            Else
                WsBCI_krga.Valcomiva = ""
            End If

            If oXml.OuterXml.Contains("<ValComNeto>") = True Then
                WsBCI_krga.Valcomneto = xObj.GetElementsByTagName("ValComNeto").Item(0).InnerText
            Else
                WsBCI_krga.Valcomneto = ""
            End If

            If oXml.OuterXml.Contains("<VlrPagar>") = True Then
                WsBCI_krga.Vlrpagar = xObj.GetElementsByTagName("VlrPagar").Item(0).InnerText
            Else
                WsBCI_krga.Vlrpagar = ""
            End If

            'campos nuevos
            If oXml.OuterXml.Contains("<CdgTraslado>") = True Then
                WsBCI_krga.Cdgtraslado = xObj.GetElementsByTagName("CdgTraslado").Item(0).InnerText
            Else
                WsBCI_krga.Cdgtraslado = ""
            End If

            If oXml.OuterXml.Contains("<IndTraslado>") = True Then
                WsBCI_krga.Indtraslado = xObj.GetElementsByTagName("IndTraslado").Item(0).InnerText
            Else
                WsBCI_krga.Indtraslado = ""
            End If

            Try

                WsBCI_krga.Firma = xObj.GetElementsByTagName("TED").Item(0).InnerXml

            Catch ex As Exception
                Console.WriteLine("Error : No se encuentra Firma en DTE")
            End Try



            Dim ConDR As Integer
            Dim DR(xObj.GetElementsByTagName("DscRcgGlobal").Count - 1) As WebReference.ZfeDscrcgglobal
            For ConDR = 0 To xObj.GetElementsByTagName("DscRcgGlobal").Count - 1
                DR(ConDR) = New WebReference.ZfeDscrcgglobal
                If oXml.OuterXml.Contains("<NroLinDR>") = True Then
                    Try
                        DR(ConDR).Nrolindr = xObj.GetElementsByTagName("NroLinDR").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Nrolindr = ""
                    End Try
                Else
                    DR(ConDR).Nrolindr = ""
                End If

                If oXml.OuterXml.Contains("<TpoMov>") = True Then
                    Try
                        DR(ConDR).Tpomov = xObj.GetElementsByTagName("TpoMov").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Tpomov = ""
                    End Try
                Else
                    DR(ConDR).Tpomov = ""
                End If

                If oXml.OuterXml.Contains("<GlosaDR>") = True Then
                    Try
                        DR(ConDR).Glosadr = xObj.GetElementsByTagName("GlosaDR").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Glosadr = ""
                    End Try
                Else
                    DR(ConDR).Glosadr = ""
                End If

                If oXml.OuterXml.Contains("<TpoValor>") = True Then
                    Try
                        DR(ConDR).Tpovalor = xObj.GetElementsByTagName("TpoValor").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Tpovalor = ""
                    End Try
                Else
                    DR(ConDR).Tpovalor = ""
                End If

                If oXml.OuterXml.Contains("<ValorDR>") = True Then
                    Try
                        DR(ConDR).Valordr = xObj.GetElementsByTagName("ValorDR").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Valordr = ""
                    End Try
                Else
                    DR(ConDR).Valordr = ""
                End If

                If oXml.OuterXml.Contains("<ValorDROtrMnda>") = True Then
                    Try
                        DR(ConDR).Valordrotrmnda = xObj.GetElementsByTagName("ValorDROtrMnda").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Valordrotrmnda = ""
                    End Try
                Else
                    DR(ConDR).Valordrotrmnda = ""
                End If

                If oXml.OuterXml.Contains("<IndExeDR>") = True Then
                    Try
                        DR(ConDR).Indexedr = xObj.GetElementsByTagName("IndExeDR").Item(ConRef).InnerText
                    Catch ex As Exception
                        DR(ConDR).Indexedr = ""
                    End Try
                Else
                    DR(ConDR).Indexedr = ""
                End If

            Next

            WsBCI_krga.Dscrcgglobal = DR

            Try
                'WsBCI_Envio.CargaDte(WsBCI_krga)
                Console.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & sObj)
                aou.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & sObj)
                aou.Flush()
            Catch ex As Exception
                aou1.WriteLine(ex.Message)
                aou1.WriteLine("ERROR, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & sObj)
                aou1.WriteLine("")
                aou1.Flush()
            End Try
            'End If

        Next


    End Sub
End Module
