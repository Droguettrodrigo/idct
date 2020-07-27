Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Collections.Generic
Imports System.Net
Imports System.Web.Services.Protocols
Imports RespuestasSII
Imports System.Configuration

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
        'RECUPERO VARIABLE PARAMETRICAS
        '
        Dim CertPath As String = My.Settings.CertPath
        Dim CertPass As String = My.Settings.CertPass
        Dim RequiereProxy As String = My.Settings.Proxy_SI_NO
        Dim ProxyServer As String = My.Settings.ProxyServer
        Dim ProxyUser As String = My.Settings.ProxyUser
        Dim ProxyPass As String = My.Settings.ProxyPass
        Dim DtePath As String = My.Settings.DtePathDteDescargados
        Dim DtePathPendientes As String = My.Settings.DtePathDtePendientes
        Dim DtePathProcOK As String = My.Settings.DtePathProcOK
        Dim maximoIntentos As String = My.Settings.maximoIntentos
        Dim DtePathProcNOK As String = My.Settings.DtePathProcNOK
        Dim VidaSemillaMil As String = My.Settings.VidaSemillaMil
        Dim DteEstOK() As String = My.Settings.DteEstOK.Split(";")
        Dim DteLogEnvio As String = My.Settings.DtePathLogEnvio
        Dim WsBCI_Carga As String = My.Settings.Ws_BCI_Carga
        Dim RUT_APLICACION As String = My.Settings.RUT_APLICACION
        Dim OMITE_VAL_SII As String = "N"
        Dim Valida_en_Sii As String = My.Settings.VALIDA_EN_SII
        Dim Omitir_Guia As String = My.Settings.OMITIR_GUIA
        Dim Reintentos_Valida_Sii As String = My.Settings.REINTENTOS_VALIDA_SII

        Dim RET_IVA_HABILITA_RZO_COMERCIAL As String = My.Settings.RET_IVA_HABILITA_RZO_COMERCIAL
        Dim RET_IVA_HABILITA_RZO_SII As String = My.Settings.RET_IVA_HABILITA_RZO_SII
        Dim RET_IVA_HABILITA_TODO As String = My.Settings.RET_IVA_HABILITA_TODO ' ESTE PARAM ES PARA HABILITAR TODO EL PROCESO DE RETENCION DE IVA

        Dim HABILITA_PROCESO_NORMAL As String = My.Settings.HABILITA_PROCESO_NORMAL ' ESTE PARAM ES PARA HABILITAR TODO EL PROCESO DE INTERCAMBIO NORMAL..SI ESTA EN N, NO SE ENVIARA MAIL NI A SAP..NO HARA NADA

        Try
            OMITE_VAL_SII = My.Settings.OMITE_VAL_SII
             
        Catch ex As Exception

        End Try

        'octu 2018
        Dim PROCESO As String
        PROCESO = Date.Now.ToString("yyyyMMddhhmmss")



        '   
        Dim rp As New RespuestasSII.RespuestasSII
        Dim resultado As String()
        'retorna
        'cantidad total
        'cantidad de procesados correctamente
        resultado = rp.Proceso(CertPath, CertPass, RequiereProxy, ProxyServer, ProxyUser, ProxyPass, DtePath, _
                               DtePathPendientes, maximoIntentos, DtePathProcNOK, VidaSemillaMil, DteEstOK, OMITE_VAL_SII, Valida_en_Sii, Omitir_Guia, Reintentos_Valida_Sii)
        Dim cantDte As Integer, cantProc As Integer
        Console.Clear()
        Console.WriteLine("*****************************************************")
        Console.WriteLine("*****************************************************")
        Console.WriteLine("*******                                       *******")
        Console.WriteLine("*****       Actualizado 01-08-2019 (SAP)        *****")
        Console.WriteLine("*******                                       *******")
        Console.WriteLine("*****************************************************")
        Console.WriteLine("*****************************************************")


        '******************************** HISTORIAL DE ACTUALIZACIONES **************************************
        '01-08-2019 Se realiza actualizacion para omitir la verificacion de la guia de despacho en el SII


        '****************************** FIN HISTORIAL DE ACTUALIZACIONES ************************************



        cantDte = resultado(0)
        cantProc = resultado(1)
        ' string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
        'GENERO NOMBRE CARPETA DIARIA
        Dim dias As String, carpetaLog As String, carpetaProc As String
        Dim fileLogExiste = False
        dias = Date.Now.ToString("yyyyMMdd")
        carpetaLog = DteLogEnvio + "\" + dias
        Dim dir1 As New DirectoryInfo(carpetaLog)
        If (dir1.Exists = False) Then
            dir1.Create()
        End If
        carpetaProc = DtePathProcOK + "\" + dias
        Dim dir2 As New DirectoryInfo(carpetaProc)
        If (dir2.Exists = False) Then
            dir2.Create()
        End If

        arrArch = Directory.GetFiles(DtePathPendientes, "*.xml")
        'leer carpeta out
        'verifico su 
        Console.WriteLine("XML para enviar a SAP: " + arrArch.Length.ToString)
        cantProc = arrArch.Length

        If cantProc > 0 Then


            'arrArch = Directory.GetFiles(DtePathProcOK, "*.xml")

            Dim aou As New StreamWriter(carpetaLog & "\Log_OK_" & DateTime.Now.ToString("ddMMyyyyhhmmss") & ".txt", True, System.Text.Encoding.GetEncoding("iso-8859-1"))
            Dim aou1 As New StreamWriter(carpetaLog & "\Log_Error_" & DateTime.Now.ToString("ddMMyyyyhhmmss") & ".txt", True, System.Text.Encoding.GetEncoding("iso-8859-1"))
            If (WsBCI_Carga.Equals("S")) Then
                aou.WriteLine("*****************************************")
                aou.WriteLine("MODO REAL: LOG OK DTE Enviado")
                aou.WriteLine("*****************************************")
                aou1.WriteLine("*****************************************")
                aou1.WriteLine("MODO REAL: LOG ERROR DTE Enviado")
                aou1.WriteLine("*****************************************")
                aou.Flush()
                aou1.Flush()
            Else
                aou.WriteLine("*****************************************")
                aou.WriteLine("MODO TEST: LOG OK DTE Enviado")
                aou.WriteLine("*****************************************")
                aou1.WriteLine("*****************************************")
                aou1.WriteLine("MODO TEST: LOG ERROR DTE Enviado")
                aou1.WriteLine("*****************************************")
                aou.Flush()
                aou1.Flush()
            End If

            'contador para xmls repetidos
            Dim cnt As Integer = 0
            For Each sObj In arrArch
                cnt += 1
                oXml = New XmlDocument()
                oXml.Load(sObj)

                'If Path.GetFileNameWithoutExtension(sObj).Split("_")(1) & Path.GetFileNameWithoutExtension(sObj).Split("_")(2) & Path.GetFileNameWithoutExtension(sObj).Split("_")(3) = "771225900520000020858" Then
                '    Console.WriteLine("")
                'End If

                'CASO RETENCION DE IVA DIC 2018 HFL

                Dim CONTINUA_PROCESO_NORMAL As Boolean = True
                If (RET_IVA_HABILITA_TODO = "S") Then

                    Try
                        Dim xmlDocRet As New XmlDocument()
                        xmlDocRet.Load(sObj)
                        xmlDocRet.PreserveWhitespace = True
                        Tipodte = xmlDocRet.GetElementsByTagName("TipoDTE")(0).InnerText
                        folio = xmlDocRet.GetElementsByTagName("Folio")(0).InnerText
                        Fchemis = xmlDocRet.GetElementsByTagName("FchEmis")(0).InnerText
                        Rznsocrecep = xmlDocRet.GetElementsByTagName("RznSocRecep")(0).InnerText
                        Rutemisor = xmlDocRet.GetElementsByTagName("RUTEmisor")(0).InnerText
                        Mnttotal = xmlDocRet.GetElementsByTagName("MntTotal")(0).InnerText
                        Mntneto = ""
                        If xmlDocRet.GetElementsByTagName("MntNeto").Count > 0 Then
                            Mntneto = xmlDocRet.GetElementsByTagName("MntNeto")(0).InnerText
                        End If
                        Iva = ""
                        If xmlDocRet.GetElementsByTagName("IVA").Count > 0 Then
                            Iva = xmlDocRet.GetElementsByTagName("IVA")(0).InnerText
                        End If
                        Rznsoc = ""
                        If xmlDocRet.GetElementsByTagName("RznSoc").Count > 0 Then
                            Rznsoc = xmlDocRet.GetElementsByTagName("RznSoc")(0).InnerText
                        End If
                        'forma de pago??
                        ''Valor 1: Contado; 2: Crédito() 3: Sin costo (entrega gratuita)
                        Dim FmaPago As String = ""
                        If xmlDocRet.GetElementsByTagName("FmaPago").Count > 0 Then
                            FmaPago = xmlDocRet.GetElementsByTagName("FmaPago")(0).InnerText
                        End If

                        Dim Retenedor As Boolean = False
                        Dim Sujeto_a_Retener As Boolean = False
                        Retenedor = esRetenedor(RUT_APLICACION)
                        Sujeto_a_Retener = esSujeto_a_Retener(Rutemisor)
                        'fechaemision se condidera para corte?
                        If (Retenedor And Sujeto_a_Retener And Tipodte = "33") Then
                            'grabo en log de sap para informar 
                            Dim sql As String = ""
                            sql = "INSERT INTO ENVIO_SAP_Y_MONITOR (proceso, [TipoDTE],[Folio],[RUTEmisor],glosa," &
                                   " archivo, codsap,glosap)"
                            sql &= "VALUES('" & PROCESO & "','" & Tipodte & "','" & folio & "','" & Rutemisor &
                                "','NO ENVIADO A SAP','" & Path.GetFileName(sObj) & "','8888','RUT SUJETO A RETENCION')"
                            Using cn As New SqlConnection(System.Configuration.ConfigurationSettings.AppSettings("BD").ToString())
                                Dim cmd As SqlCommand = New SqlCommand
                                With cmd
                                    .Connection = cn
                                    .Connection.Open()
                                    .CommandText = sql
                                    .CommandType = CommandType.Text
                                End With
                                cmd.ExecuteNonQuery()
                                cn.Close()
                            End Using
                            'PROCESO SE RETENCION-RECHAZO-AVISO-GRABACION BD
                            Dim rs As Boolean
                            rs = procesoRetencion(Rutemisor, Tipodte, folio, Fchemis, Rznsoc, Mntneto, Iva, Mnttotal, xmlDocRet, System.Configuration.ConfigurationSettings.AppSettings("BD").ToString(), RET_IVA_HABILITA_RZO_COMERCIAL, RET_IVA_HABILITA_RZO_SII
                                                  )
                            'muevoarchivo a nueva ruta de archivos rechazados por sujeto iva
                            Dim myFile As New FileInfo(sObj)
                            If Directory.Exists(carpetaProc & "\RETENCION") = False Then
                                Directory.CreateDirectory(carpetaProc & "\RETENCION")
                            End If

                            If File.Exists(carpetaProc & "\RETENCION\" & Path.GetFileName(sObj)) Then
                                File.Move(sObj, carpetaProc & "\RETENCION\" & Path.GetFileName(sObj))
                                Console.WriteLine("DTE Enviado A Retencion, " & "Folio : " & folio & " " & Path.GetFileName(sObj))
                                aou.WriteLine("DTE Enviado A Retencion, " & "Folio : " & folio & " " & Path.GetFileName(sObj))
                            Else
                                File.Move(sObj, carpetaProc & "\RETENCION\" & myFile.Name)
                                Console.WriteLine("DTE Enviado A Retencion, " & "Folio : " & folio & " " & myFile.FullName)
                                aou.WriteLine("DTE Enviado A Retencion, " & "Folio : " & folio & " " & myFile.FullName)

                            End If
                            CONTINUA_PROCESO_NORMAL = False

                        Else
                            'continua normal
                            CONTINUA_PROCESO_NORMAL = True

                        End If

                    Catch ex As Exception

                    End Try
                End If

                'fin caso retencion de iva    

                If (HABILITA_PROCESO_NORMAL = "N") Then   'deshabilito proceso normal...ideal para pruebas
                    CONTINUA_PROCESO_NORMAL = False
                    Console.WriteLine("EL PROCESO DE INTERCAMBIO ESTA DESHABILITADO, NO SE GRABARA, NI PROCESARA, NI ENVIARA A SAP. SI QUIERES HABILITAR CAMBIA EL PARAMETRO: HABILITA_PROCESO_NORMAL: S")


                End If

                If CONTINUA_PROCESO_NORMAL = True Then

                    xObj = oXml.GetElementsByTagName("Documento").Item(0)
                    If xObj Is Nothing Then
                        xObj = oXml.GetElementsByTagName("Liquidacion").Item(0)
                    End If


                    Dim WsBCI_krga As New WebReference.CargaDte
                    'Dim WsBCI_Envio As New WebReference.zservicios_fe_v1
                    'Dim WsBCI_Envio As New WebReference.zservicios_fe_v2
                    Dim WsBCI_Envio As New WebReference.zservicios_fe_v2


                    Dim WsBCI_Detalle As New WebReference.ZfeDetalle
                    Dim WsBCI_Impuesto As New WebReference.ZfeImpuestos
                    Dim WsBCI_Referencia As New WebReference.ZfeReferencia

                    Dim objCredential As New System.Net.NetworkCredential()


                    'objCredential.UserName = "ws_fe"

                    'objCredential.Password = "INICIO17"

                    objCredential.UserName = My.Settings.WS_BCI_USER

                    objCredential.Password = My.Settings.WS_BCI_PASS



                    WsBCI_Envio.Credentials = objCredential

                    'Dim det(xObj.GetElementsByTagName("Detalle").Count - 1) As WebReference.ZfeDetalle
                    'det(xObj.GetElementsByTagName("Detalle").Count - 1) = New WebReference.ZfeDetalle

                    Dim heshem() As String = Nothing
                    Dim ValHesHem As String = ""


                    Dim det(xObj.GetElementsByTagName("Detalle").Count - 1) As WebReference.ZfeDetalle
                    Dim ConDet As Integer = 0
                    Dim newRef As Integer = 0
                    For ConDet = 0 To xObj.GetElementsByTagName("Detalle").Count - 1
                        det(ConDet) = New WebReference.ZfeDetalle
                        If oXml.OuterXml.Contains("DscItem") = True Then
                            Try
                                det(ConDet).Dscitem = xObj.GetElementsByTagName("DscItem").Item(ConDet).InnerText
                                heshem = det(ConDet).Dscitem.Split(" ")
                                Dim cHes As Integer
                                For cHes = 0 To heshem.Length - 1
                                    If heshem(cHes).Length >= 10 Then
                                        If heshem(cHes).Substring(0, 4) = "1000" Or heshem(cHes).Substring(0, 4) = "5000" Then
                                            'If heshem(cHes).Substring(0, 10) = "1000" Or heshem(cHes).Substring(0, 3) = "5000" Then
                                            ValHesHem = heshem(cHes).Substring(0, 10)
                                            newRef = 1
                                        End If
                                    End If

                                Next

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
                                det(ConDet).Montoitem = "0"
                            End Try
                        Else
                            det(ConDet).Montoitem = "0"
                        End If

                        If oXml.OuterXml.Contains("<NmbItem>") = True Then
                            Try
                                det(ConDet).Nmbitem = xObj.GetElementsByTagName("NmbItem").Item(ConDet).InnerText

                                heshem = det(ConDet).Nmbitem.Split(" ")
                                Dim cHes As Integer
                                For cHes = 0 To heshem.Length - 1
                                    If heshem(cHes).Length >= 10 Then
                                        If heshem(cHes).Substring(0, 4) = "1000" Or heshem(cHes).Substring(0, 4) = "5000" Then
                                            'If heshem(cHes).Substring(0, 10) = "1000" Or heshem(cHes).Substring(0, 3) = "5000" Then
                                            ValHesHem = heshem(cHes).Substring(0, 10)
                                            newRef = 1
                                        End If
                                    End If

                                Next

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
                    Dim Ref(xObj.GetElementsByTagName("Referencia").Count + newRef - 1) As WebReference.ZfeReferencia
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
                                If Ref(ConRef).Tpodocref.ToUpper = "HES" Or Ref(ConRef).Tpodocref.ToUpper = "HEM" Then
                                    ValHesHem = ""
                                End If

                            Catch ex As Exception
                                Ref(ConRef).Tpodocref = ""
                            End Try
                        Else
                            Ref(ConRef).Tpodocref = ""
                        End If
                    Next

                    'referencia hes en detalle

                    If xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText = "33" Or xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText = "34" Or xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText = "46" Then
                        If ValHesHem <> "" Then
                            Ref(ConRef) = New WebReference.ZfeReferencia
                            Dim tipodatoheshem As String = ""

                            If ValHesHem.Substring(0, 4) = "1000" Then
                                tipodatoheshem = "HES"
                            Else
                                tipodatoheshem = "HEM"
                            End If

                            Ref(ConRef).Nrolinref = ConRef + 1
                            Ref(ConRef).Tpodocref = tipodatoheshem
                            Ref(ConRef).Folioref = ValHesHem
                            Ref(ConRef).Fchref = ""
                            Ref(ConRef).Razonref = ""

                        End If
                        'If ValHesHem <> "" Then

                        '    Dim ConRef1 As Integer
                        '    Dim Ref1(0) As WebReference.ZfeReferencia
                        '    Dim tipodatoheshem As String = ""
                        '    For ConRef1 = 0 To 1
                        '        Ref1(ConRef1) = New WebReference.ZfeReferencia

                        '        If ValHesHem.Substring(0, 4) = "1000" Then
                        '            tipodatoheshem = "HES"
                        '        Else
                        '            tipodatoheshem = "HEM"
                        '        End If


                        '        Ref1(0).Nrolinref = 1
                        '        Ref1(0).Tpodocref = tipodatoheshem
                        '        Ref1(0).Folioref = ValHesHem
                        '        Ref1(0).Fchref = ""
                        '        Ref1(0).Razonref = ""
                        '        Ref = Ref1
                        '        Exit For
                        '    Next


                        'End If
                    End If
                    ValHesHem = ""

                    WsBCI_krga.Referencia = Ref

                    'Activar WS cargadte
                    Try

                        If oXml.OuterXml.Contains("<CdgIntRecep>") = True Then
                            WsBCI_krga.Cdgintrecep = xObj.GetElementsByTagName("CdgIntRecep").Item(0).InnerText
                        Else
                            WsBCI_krga.Cdgintrecep = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Cdgintrecep = ""
                    End Try


                    Try
                        If oXml.OuterXml.Contains("<Cdgvendedor>") = True Then
                            WsBCI_krga.Cdgvendedor = xObj.GetElementsByTagName("Cdgvendedor").Item(0).InnerText
                        Else
                            WsBCI_krga.Cdgvendedor = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Cdgvendedor = ""
                    End Try
                    Try
                        If oXml.OuterXml.Contains("<CiudadOrigen>") = True Then
                            WsBCI_krga.Ciudadorigen = xObj.GetElementsByTagName("CiudadOrigen").Item(0).InnerText
                        Else
                            WsBCI_krga.Ciudadorigen = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Ciudadorigen = ""
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<CiudadRecep>") = True Then
                            WsBCI_krga.Ciudadrecep = xObj.GetElementsByTagName("CiudadRecep").Item(0).InnerText
                        Else
                            WsBCI_krga.Ciudadrecep = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Ciudadrecep = ""
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<CmnaOrigen>") = True Then
                            WsBCI_krga.Cmnaorigen = xObj.GetElementsByTagName("CmnaOrigen").Item(0).InnerText
                        Else
                            WsBCI_krga.Cmnaorigen = "COMUNA ORIGEN"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Cmnaorigen = "COMUNA ORIGEN"
                    End Try

                    If oXml.OuterXml.Contains("<CmnaRecep>") = True Then
                        WsBCI_krga.Cmnarecep = xObj.GetElementsByTagName("CmnaRecep").Item(0).InnerText
                    Else
                        WsBCI_krga.Cmnarecep = "COMUNA RECEPTOR"
                    End If

                    Try
                        If oXml.OuterXml.Contains("<CorreoEmisor>") = True Then
                            WsBCI_krga.Correoemisor = xObj.GetElementsByTagName("CorreoEmisor").Item(0).InnerText
                        Else
                            WsBCI_krga.Correoemisor = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Correoemisor = ""
                    End Try

                    Try

                        If oXml.OuterXml.Contains("<CorreoRecep>") = True Then
                            WsBCI_krga.Correorecep = xObj.GetElementsByTagName("CorreoRecep").Item(0).InnerText
                        Else
                            WsBCI_krga.Correorecep = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Correorecep = ""

                    End Try

                    Try

                        If oXml.OuterXml.Contains("<DirOrigen>") = True Then
                            If xObj.GetElementsByTagName("DirOrigen").Item(0).InnerText.Length < 60 Then
                                WsBCI_krga.Dirorigen = xObj.GetElementsByTagName("DirOrigen").Item(0).InnerText
                            Else
                                WsBCI_krga.Dirorigen = xObj.GetElementsByTagName("DirOrigen").Item(0).InnerText.Substring(0, 60)
                            End If

                        Else
                            WsBCI_krga.Dirorigen = "DIRECCION DE ORIGEN"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Dirorigen = "DIRECCION DE ORIGEN"
                    End Try

                    If oXml.OuterXml.Contains("<DirRecep>") = True Then
                        WsBCI_krga.Dirrecep = xObj.GetElementsByTagName("DirRecep").Item(0).InnerText
                    Else
                        WsBCI_krga.Dirrecep = "DIRECCION RECEPTOR"
                    End If

                    If oXml.OuterXml.Contains("<FchEmis>") = True Then
                        WsBCI_krga.Fchemis = xObj.GetElementsByTagName("FchEmis").Item(0).InnerText
                    Else
                        WsBCI_krga.Fchemis = "AAAA-MM-DD"
                    End If

                    Try
                        If oXml.OuterXml.Contains("<FchVenc>") = True Then
                            WsBCI_krga.Fchvenc = xObj.GetElementsByTagName("FchVenc").Item(0).InnerText
                        Else
                            WsBCI_krga.Fchvenc = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Fchvenc = ""

                    End Try

                    Try
                        If oXml.OuterXml.Contains("<FmaPago>") = True Then
                            WsBCI_krga.Fmapago = xObj.GetElementsByTagName("FmaPago").Item(0).InnerText
                        Else
                            WsBCI_krga.Fmapago = ""
                        End If

                    Catch ex As Exception
                        WsBCI_krga.Fmapago = ""

                    End Try

                    If oXml.OuterXml.Contains("<Folio>") = True Then
                        WsBCI_krga.Folio = xObj.GetElementsByTagName("Folio").Item(0).InnerText
                    Else
                        WsBCI_krga.Folio = "0"
                    End If

                    If oXml.OuterXml.Contains("<GiroEmis>") = True Then
                        WsBCI_krga.Giroemis = xObj.GetElementsByTagName("GiroEmis").Item(0).InnerText.Trim
                    Else
                        WsBCI_krga.Giroemis = "GIRO DEL NEGOCIO DEL EMISOR"
                    End If

                    Try
                        If oXml.OuterXml.Contains("<GiroRecep>") = True Then
                            WsBCI_krga.Girorecep = xObj.GetElementsByTagName("GiroRecep").Item(0).InnerText
                        Else
                            WsBCI_krga.Girorecep = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Girorecep = ""
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<IVA>") = True Then
                            WsBCI_krga.Iva = Trim(xObj.GetElementsByTagName("IVA").Item(0).InnerText)
                        Else
                            WsBCI_krga.Iva = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Iva = "0"
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<IVANoRet>") = True Then
                            WsBCI_krga.Ivanoret = xObj.GetElementsByTagName("IVANoRet").Item(0).InnerText
                        Else
                            WsBCI_krga.Ivanoret = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Ivanoret = "0"
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<MntExe>") = True Then
                            WsBCI_krga.Mntexe = xObj.GetElementsByTagName("MntExe").Item(0).InnerText
                        Else
                            WsBCI_krga.Mntexe = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Mntexe = "0"
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<MntNeto>") = True Then
                            WsBCI_krga.Mntneto = Trim(xObj.GetElementsByTagName("MntNeto").Item(0).InnerText)
                        Else
                            WsBCI_krga.Mntneto = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Mntneto = "0"
                    End Try
                    If oXml.OuterXml.Contains("<MntTotal>") = True Then
                        Dim MtoTotal As Long
                        If xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText = "46" Then
                            Try
                                MtoTotal = CLng(WsBCI_krga.Mntneto) + CLng(WsBCI_krga.Iva)
                            Catch ex As Exception
                                MtoTotal = "0"
                            End Try
                        Else
                            MtoTotal = Trim(xObj.GetElementsByTagName("MntTotal").Item(0).InnerText)
                        End If
                        WsBCI_krga.Mnttotal = MtoTotal
                    Else
                        WsBCI_krga.Mnttotal = "0"
                    End If
                    Try
                        If oXml.OuterXml.Contains("<MontoNF>") = True Then
                            WsBCI_krga.Montonf = xObj.GetElementsByTagName("MontoNF").Item(0).InnerText
                        Else
                            WsBCI_krga.Montonf = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Montonf = ""
                    End Try

                    Try
                        If oXml.OuterXml.Contains("<MontoPeriodo>") = True Then
                            WsBCI_krga.Montoperiodo = xObj.GetElementsByTagName("MontoPeriodo").Item(0).InnerText
                        Else
                            WsBCI_krga.Montoperiodo = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Montoperiodo = ""
                    End Try

                    If oXml.OuterXml.Contains("<RUTEmisor>") = True Then
                        WsBCI_krga.Rutemisor = xObj.GetElementsByTagName("RUTEmisor").Item(0).InnerText
                    Else
                        WsBCI_krga.Rutemisor = "99999999-9"
                    End If

                    If oXml.OuterXml.Contains("<RUTRecep>") = True Then
                        WsBCI_krga.Rutrecep = xObj.GetElementsByTagName("RUTRecep").Item(0).InnerText
                    Else
                        WsBCI_krga.Rutrecep = "99999999-9"
                    End If

                    If oXml.OuterXml.Contains("<RznSoc>") = True Then
                        WsBCI_krga.Rznsoc = xObj.GetElementsByTagName("RznSoc").Item(0).InnerText
                    Else
                        WsBCI_krga.Rznsoc = "NOMBRE O RAZON SOCIAL EMISOR"
                    End If

                    If oXml.OuterXml.Contains("<RznSocRecep>") = True Then
                        WsBCI_krga.Rznsocrecep = xObj.GetElementsByTagName("RznSocRecep").Item(0).InnerText
                    Else
                        WsBCI_krga.Rznsocrecep = "NOMBRE O RAZON SOCIAL RECEPTOR"
                    End If

                    Try
                        If oXml.OuterXml.Contains("<SaldoAnterior>") = True Then
                            WsBCI_krga.Saldoanterior = xObj.GetElementsByTagName("SaldoAnterior").Item(0).InnerText
                        Else
                            WsBCI_krga.Saldoanterior = ""
                        End If

                    Catch ex As Exception
                        WsBCI_krga.Saldoanterior = ""
                    End Try


                    Try
                        If oXml.OuterXml.Contains("<TasaIVA>") = True Then
                            Dim _tasaiva = xObj.GetElementsByTagName("TasaIVA").Item(0).InnerText
                            If (xObj.GetElementsByTagName("TasaIVA").Item(0).InnerText = "019.00") Then
                                _tasaiva = "19"

                            End If

                            WsBCI_krga.Tasaiva = _tasaiva
                        Else
                            WsBCI_krga.Tasaiva = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Tasaiva = "0"
                    End Try

                    If oXml.OuterXml.Contains("<TipoDTE>") = True Then
                        WsBCI_krga.Tipodte = xObj.GetElementsByTagName("TipoDTE").Item(0).InnerText
                    Else
                        WsBCI_krga.Tipodte = "0"
                    End If

                    Try
                        If oXml.OuterXml.Contains("<TpoMoneda>") = True Then
                            WsBCI_krga.Tpomoneda = xObj.GetElementsByTagName("TpoMoneda").Item(0).InnerText
                        Else
                            WsBCI_krga.Tpomoneda = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Tpomoneda = ""
                    End Try


                    Try
                        If oXml.OuterXml.Contains("<ValComExe>") = True Then
                            WsBCI_krga.Valcomexe = xObj.GetElementsByTagName("ValComExe").Item(0).InnerText
                        Else
                            WsBCI_krga.Valcomexe = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Valcomexe = "0"

                    End Try

                    Try
                        If oXml.OuterXml.Contains("<ValComIVA>") = True Then
                            WsBCI_krga.Valcomiva = xObj.GetElementsByTagName("ValComIVA").Item(0).InnerText
                        Else
                            WsBCI_krga.Valcomiva = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Valcomiva = ""

                    End Try

                    Try

                        If oXml.OuterXml.Contains("<ValComNeto>") = True Then
                            WsBCI_krga.Valcomneto = xObj.GetElementsByTagName("ValComNeto").Item(0).InnerText
                        Else
                            WsBCI_krga.Valcomneto = "0"
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Valcomneto = "0"
                    End Try
                    Try
                        If oXml.OuterXml.Contains("<VlrPagar>") = True Then
                            WsBCI_krga.Vlrpagar = xObj.GetElementsByTagName("VlrPagar").Item(0).InnerText
                        Else
                            WsBCI_krga.Vlrpagar = ""
                        End If
                    Catch ex As Exception
                        WsBCI_krga.Vlrpagar = ""
                    End Try
                    'campos nuevos omitido 10102018
                    'Try

                    '    If oXml.OuterXml.Contains("<CdgTraslado>") = True Then
                    '        WsBCI_krga.cdCdgtraslado = xObj.GetElementsByTagName("CdgTraslado").Item(0).InnerText
                    '    Else
                    '        WsBCI_krga.Cdgtraslado = ""
                    '    End If
                    'Catch ex As Exception
                    '    WsBCI_krga.Cdgtraslado = ""

                    'End Try

                    'Try
                    '    If oXml.OuterXml.Contains("<IndTraslado>") = True Then
                    '        WsBCI_krga.Indtraslado = xObj.GetElementsByTagName("IndTraslado").Item(0).InnerText
                    '    Else
                    '        WsBCI_krga.Indtraslado = ""
                    '    End If
                    'Catch ex As Exception
                    '    WsBCI_krga.Indtraslado = ""
                    'End Try

                    'Try

                    '    WsBCI_krga.Firma = xObj.GetElementsByTagName("TED").Item(0).InnerXml

                    'Catch ex As Exception
                    '    Console.WriteLine("Error : No se encuentra Firma en DTE")
                    'End Try



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
                        If (WsBCI_Carga.Equals("S")) Then


                            'validacion
                            If ((WsBCI_krga.Rutrecep = RUT_APLICACION) Or (WsBCI_krga.Tipodte = "46" And WsBCI_krga.Rutemisor = RUT_APLICACION)) Then





                                Dim RSP As New WebReference.CargaDteResponse

                                Console.WriteLine("Enviando a sap: " & WsBCI_krga.Rutemisor & " - " & WsBCI_krga.Tipodte & " - " & WsBCI_krga.Folio & "....")


                                RSP = WsBCI_Envio.CargaDte(WsBCI_krga)

                                Dim codsap, glosap As String
                                Console.WriteLine("Envio OK.")
                                Try
                                    Console.WriteLine("Respuesta SAP OK.")
                                    Console.WriteLine("Codigoretorno: " & RSP.Codigoretorno.ToString())
                                    Console.WriteLine("Descripcioncodigo: " & RSP.Descripcioncodigo.ToString())
                                    codsap = RSP.Codigoretorno.ToString()
                                    glosap = RSP.Descripcioncodigo.ToString()

                                Catch ex As Exception
                                    Console.WriteLine("Respuesta SAP ERROR: " & ex.Message.ToString())
                                    codsap = "ERROR"
                                    glosap = ex.Message.ToString()
                                End Try




                                Dim conta = 0
                                Dim sqlCnx As New SqlConnection()
                                Dim sqlCmd As New SqlCommand()

                                sqlCnx.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("BD").ToString()
                                sqlCmd.Connection = sqlCnx

                                sqlCmd.CommandTimeout = 999999999
                                sqlCmd.CommandType = CommandType.Text
                                sqlCnx.Open()

                                Dim xmlDoc As New XmlDocument()
                                xmlDoc.Load(sObj)

                                Dim xmlObject As New XmlDocument()
                                xmlDoc.PreserveWhitespace = True
                                Tipodte = xmlDoc.GetElementsByTagName("TipoDTE")(0).InnerText
                                folio = xmlDoc.GetElementsByTagName("Folio")(0).InnerText
                                Fchemis = xmlDoc.GetElementsByTagName("FchEmis")(0).InnerText
                                Rutemisor = xmlDoc.GetElementsByTagName("RUTEmisor")(0).InnerText


                                Rznsoc = ""
                                If xmlDoc.GetElementsByTagName("RznSoc").Count > 0 Then
                                    Rznsoc = xmlDoc.GetElementsByTagName("RznSoc")(0).InnerText
                                End If

                                Giroemis = ""
                                If xmlDoc.GetElementsByTagName("GiroEmis").Count > 0 Then
                                    Giroemis = xmlDoc.GetElementsByTagName("GiroEmis")(0).InnerText
                                End If

                                Dim Acteco As String = ""
                                If xmlDoc.GetElementsByTagName("Acteco").Count > 0 Then
                                    Acteco = xmlDoc.GetElementsByTagName("Acteco")(0).InnerText
                                End If

                                Dim CdgSIISucur As String = ""
                                If xmlDoc.GetElementsByTagName("CdgSIISucur").Count > 0 Then
                                    CdgSIISucur = xmlDoc.GetElementsByTagName("CdgSIISucur")(0).InnerText
                                End If
                                Dirorigen = ""
                                If xmlDoc.GetElementsByTagName("DirOrigen").Count > 0 Then
                                    Dirorigen = xmlDoc.GetElementsByTagName("DirOrigen")(0).InnerText.Replace("'", "''")
                                End If
                                Cmnaorigen = ""
                                If xmlDoc.GetElementsByTagName("Cmnaorigen").Count > 0 Then

                                    Cmnaorigen = xmlDoc.GetElementsByTagName("CmnaOrigen")(0).InnerText
                                End If

                                Ciudadorigen = ""
                                If xmlDoc.GetElementsByTagName("CiudadOrigen").Count > 0 Then
                                    Ciudadorigen = xmlDoc.GetElementsByTagName("CiudadOrigen")(0).InnerText
                                End If


                                Mntneto = "0"
                                If xmlDoc.GetElementsByTagName("MntNeto").Count > 0 Then
                                    Mntneto = xmlDoc.GetElementsByTagName("MntNeto")(0).InnerText
                                End If

                                Mntexe = "0"
                                If xmlDoc.GetElementsByTagName("MntExe").Count > 0 Then
                                    Mntexe = xmlDoc.GetElementsByTagName("MntExe")(0).InnerText
                                End If


                                Tasaiva = "0"
                                If xmlDoc.GetElementsByTagName("TasaIVA").Count > 0 Then
                                    Tasaiva = xmlDoc.GetElementsByTagName("TasaIVA")(0).InnerText
                                End If


                                Iva = "0"
                                If xmlDoc.GetElementsByTagName("IVA").Count > 0 Then
                                    Iva = xmlDoc.GetElementsByTagName("IVA")(0).InnerText
                                End If

                                Mnttotal = xmlDoc.GetElementsByTagName("MntTotal")(0).InnerText
                                Dim archivo As String = ""
                                Dim fecRegistro As String = ""
                                Dim estadoAprobComer As String = ""
                                Dim glosaAprobComer As String = ""
                                Dim fechaAprobComer As String = ""
                                Dim fecRecepcion As String = ""
                                Dim estadoRecepcion As String = ""
                                Dim glosaRecepcion As String = ""

                                Dim sql As String = ""

                                conta += 1
                                'grabo new log
                                sql = "INSERT INTO ENVIO_SAP_Y_MONITOR (proceso, [TipoDTE],[Folio],[RUTEmisor],glosa,archivo, codsap,glosap)"
                                sql &= "VALUES('" & PROCESO & "','" & Tipodte & "','" & folio & "','" & Rutemisor & "','ENVIADO A SAP','" & Path.GetFileName(sObj) & "','" & codsap & "','" & glosap & "')"
                                sqlCmd.CommandText = sql
                                sqlCmd.ExecuteNonQuery()


                                'graba a bd del monitor
                                Dim estado_graba_monitor = "N"
                                Dim glosa_graba_monitor = ""
                                Try
                                    sqlCmd.CommandText = "select COUNT(*) from [dte_recibidos_EMAIL] where Folio='" + folio + "' and TipoDTE='" + Tipodte + "' and RUTEmisor='" + Rutemisor + "'"
                                    Dim rcntG As Object = sqlCmd.ExecuteScalar()
                                    Dim cc As String = rcntG.ToString()
                                    If (cc = 0) Then

                                        sql = "INSERT INTO [dte_recibidos_EMAIL]([id],[TipoDTE],[Folio],[FchEmis],[RUTEmisor],[RznSoc],[GiroEmis],[Acteco],[CdgSIISucur],[DirOrigen],[CmnaOrigen],[CiudadOrigen],[MntNeto],[MntExe],[TasaIVA],[IVA],[MntTotal],[archivo],[fecRegistro],[estadoAprobComer],[glosaAprobComer],[fechaAprobComer],[fecRecepcion],[estadoRecepcion],[glosaRecepcion])"
                                        sql &= "VALUES('" & conta & "','" & Tipodte & "','" & folio & "','" & Fchemis & "','" & Rutemisor & "','" & Rznsoc.Replace("'", "''") & "','" & Giroemis.Replace("'", "''") & "','" & Acteco & "','" & CdgSIISucur & "','" & Dirorigen.Replace("'", "''") & "','" & Cmnaorigen.Replace("'", "''") & "','" & Ciudadorigen.Replace("'", "''") & "','" & Mntneto & "','" & Mntexe & "','" & Tasaiva & "','" & Iva & "','" & Mnttotal & "','" & xmlDoc.OuterXml.Replace("'", "''") & "'," & "GETDATE()" & ",'" & estadoAprobComer & "','" & glosaAprobComer & "','" & fechaAprobComer & "','" & fecRecepcion & "','" & estadoRecepcion & "','" & glosaRecepcion & "')"
                                        sqlCmd.CommandText = sql
                                        sqlCmd.ExecuteNonQuery()

                                        'grabo new log
                                        sql = "INSERT INTO ENVIO_SAP_Y_MONITOR (proceso,[TipoDTE],[Folio],[RUTEmisor],glosa,archivo)"
                                        sql &= "VALUES('" & PROCESO & "','" & Tipodte & "','" & folio & "','" & Rutemisor & "','ENVIADO A MONITOR','" & Path.GetFileName(sObj) & "')"
                                        sqlCmd.CommandText = sql
                                        sqlCmd.ExecuteNonQuery()


                                        estado_graba_monitor = "S"
                                    End If
                                    'grabo historial de las veces que llega
                                    sql = "INSERT INTO [dte_recibidos_EMAIL_HISTORIAL]([id],[TipoDTE],[Folio],[FchEmis],[RUTEmisor],[RznSoc],[GiroEmis],[Acteco],[CdgSIISucur],[DirOrigen],[CmnaOrigen],[CiudadOrigen],[MntNeto],[MntExe],[TasaIVA],[IVA],[MntTotal],[archivo],[fecRegistro],[estadoAprobComer],[glosaAprobComer],[fechaAprobComer],[fecRecepcion],[estadoRecepcion],[glosaRecepcion])"
                                    sql &= "VALUES('" & conta & "','" & Tipodte & "','" & folio & "','" & Fchemis & "','" & Rutemisor & "','" & Rznsoc.Replace("'", "''") & "','" & Giroemis.Replace("'", "''") & "','" & Acteco & "','" & CdgSIISucur & "','" & Dirorigen.Replace("'", "''") & "','" & Cmnaorigen.Replace("'", "''") & "','" & Ciudadorigen.Replace("'", "''") & "','" & Mntneto & "','" & Mntexe & "','" & Tasaiva & "','" & Iva & "','" & Mnttotal & "','" & xmlDoc.OuterXml.Replace("'", "''") & "'," & "GETDATE()" & ",'" & estadoAprobComer & "','" & glosaAprobComer & "','" & fechaAprobComer & "','" & fecRecepcion & "','" & estadoRecepcion & "','" & glosaRecepcion & "')"
                                    sqlCmd.CommandText = sql
                                    sqlCmd.ExecuteNonQuery()
                                    sqlCnx.Close()

                                Catch ex As Exception
                                    estado_graba_monitor = "N"
                                    glosa_graba_monitor = "ERROR: " + ex.Message

                                End Try

                                'Console.WriteLine("grabado en BD")

                            Else 'En el caso que el receptor no corresponda al rut de la aplicación
                                Dim sql As String = ""
                                sql = "INSERT INTO ENVIO_SAP_Y_MONITOR (proceso, [TipoDTE],[Folio],[RUTEmisor],glosa,archivo, codsap,glosap)"
                                sql &= "VALUES('" & PROCESO & "','" & Tipodte & "','" & folio & "','" & Rutemisor & "','NO ENVIADO A SAP','" & Path.GetFileName(sObj) & "','9999','Rut Receptor Inválido')"
                                Using cn As New SqlConnection(System.Configuration.ConfigurationSettings.AppSettings("BD").ToString())
                                    Dim cmd As SqlCommand = New SqlCommand
                                    With cmd
                                        .Connection = cn
                                        .Connection.Open()
                                        .CommandText = sql
                                        .CommandType = CommandType.Text
                                    End With
                                    cmd.ExecuteNonQuery()
                                    cn.Close()
                                End Using
                                Console.WriteLine("DTE NO VÁLIDO (RUT RECEPTOR INVÁLIDO), " & "Folio : " & WsBCI_krga.Folio.ToString & " " & oXml.Name)
                                aou.WriteLine("DTE NO VÁLIDO (RUT RECEPTOR INVÁLIDO), " & "Folio : " & WsBCI_krga.Folio.ToString & " " & oXml.Name)
                            End If
                            'Console.WriteLine("")
                        Else

                            'xml rut no BCI

                        End If
                        'muevo archivo a destino final

                        'File.Delete(sObj)

                        Dim myFile As New FileInfo(sObj)


                        If File.Exists(carpetaProc & "\" & Path.GetFileName(sObj)) Then
                            File.Move(sObj, carpetaProc & "\" & Path.GetFileName(sObj))
                            Console.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & Path.GetFileName(sObj))
                            aou.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & Path.GetFileName(sObj))
                        Else
                            File.Move(sObj, carpetaProc & "\" & myFile.Name)
                            Console.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & myFile.FullName)
                            aou.WriteLine("DTE Enviado, " & "Folio : " & WsBCI_krga.Folio.ToString & " " & myFile.FullName)

                        End If

                        aou.Flush()
                    Catch ex As Exception
                        aou1.WriteLine(ex.Message)
                        aou1.WriteLine("ERROR, (" & Date.Today & " Folio : " & WsBCI_krga.Folio.ToString & " " & sObj)
                        aou1.WriteLine("")
                        aou1.Flush()
                        Console.WriteLine("ERROR, (" & Date.Today & " Folio : " & WsBCI_krga.Folio.ToString & " " & sObj)

                        sqlCnx.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("BD").ToString()
                        sqlCmd.Connection = sqlCnx

                        sqlCmd.CommandTimeout = 999999999
                        sqlCmd.CommandType = CommandType.Text
                        sqlCnx.Open()
                        'grabo new log
                        Dim sQL As String
                        sQL = "INSERT INTO ENVIO_SAP_Y_MONITOR (proceso,[TipoDTE],[Folio],[RUTEmisor],glosa,archivo)"
                        sQL &= "VALUES('" & PROCESO & "','" & WsBCI_krga.Tipodte.ToString & "','" & WsBCI_krga.Folio.ToString & "','" & WsBCI_krga.Rutemisor.ToString & "','ERROR SAP - " & ex.Message & "','" & Path.GetFileName(sObj) & "')"
                        sqlCmd.CommandText = sQL
                        sqlCmd.ExecuteNonQuery()
                        sqlCnx.Close()


                    End Try

                    'proceso normal
                End If



            Next

        End If

    End Sub

    Private Function esRetenedor(ByVal rut As String) As Boolean
        Dim rsp = False
        Try

            Dim sqlCnx As New SqlConnection()
            Dim sqlCmd As New SqlCommand()

            sqlCnx.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("BD").ToString()
            sqlCmd.Connection = sqlCnx

            sqlCmd.CommandTimeout = 999999999
            sqlCmd.CommandType = CommandType.Text
            sqlCnx.Open()

            sqlCmd.CommandText = "select COUNT(*) from RET_IVA_AGENTES_RETENEDORES where RUT='" + rut + "'  "
            Dim rcntG As Object = sqlCmd.ExecuteScalar()
            Dim cc As String = rcntG.ToString()
            If (cc <> "0") Then
                rsp = True

            End If
            sqlCnx.Close()
        Catch ex As Exception

        End Try


        Return rsp
    End Function
    Private Function esSujeto_a_Retener(ByVal rut As String) As Boolean
        Dim rsp = False
        Try

            Dim sqlCnx As New SqlConnection()
            Dim sqlCmd As New SqlCommand()

            sqlCnx.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("BD").ToString()
            sqlCmd.Connection = sqlCnx

            sqlCmd.CommandTimeout = 999999999
            sqlCmd.CommandType = CommandType.Text
            sqlCnx.Open()

            sqlCmd.CommandText = "select COUNT(*) from RET_IVA_SUJETOS_A_RETENER where RUT='" + rut + "'  "
            Dim rcntG As Object = sqlCmd.ExecuteScalar()
            Dim cc As String = rcntG.ToString()
            If (cc <> "0") Then
                rsp = True

            End If

            sqlCnx.Close()
        Catch ex As Exception

        End Try


        Return rsp
    End Function


    Private Function procesoRetencion(ByVal rut As String, ByVal tipo As String, ByVal folio As String, ByVal fecemi As String,
                                      ByVal nombre As String, ByVal neto As String, ByVal iva As String, ByVal total As String, ByVal xml As XmlDocument, ByVal conexionSQL As String, ByVal RET_IVA_HABILITA_RZO_COMERCIAL As String, ByVal RET_IVA_HABILITA_RZO_SII As String) As Boolean
        Dim zona As New RespuestasSII.RespuestasSII
        Dim rsp = False

        'graba a tabla disponible a generar
        Dim ticket As String
        ticket = zona.grabaPendienteRetener(rut, tipo, folio, fecemi, nombre, neto, iva, total, xml, conexionSQL)

        If (ticket <> "") Then
            'rechazo comercial
            Dim r1 As Boolean
            If (RET_IVA_HABILITA_RZO_COMERCIAL = "S") Then
                r1 = zona.rechazoComercial(rut, tipo, folio, fecemi, total, nombre, My.Settings.CertPath, My.Settings.CertPass, conexionSQL, My.Settings.RUTA_TMP_RET_IVA)
            End If


            'rechazo en sii
            Dim r2 As Boolean
            If (RET_IVA_HABILITA_RZO_SII = "S") Then
                r2 = zona.rechazoSII(rut, tipo, folio, fecemi, total, My.Settings.CertPath, My.Settings.CertPass, conexionSQL)
            End If
            'envio alertas de retencion
            Dim r4 As Boolean
            r4 = zona.enviaAlertaRetencion(rut, tipo, folio, fecemi, total, xml, conexionSQL, My.Settings.RUTA_TMP_RET_IVA, ticket, nombre)


            rsp = True
        Else
            rsp = False

        End If

        Return rsp
    End Function

End Module
