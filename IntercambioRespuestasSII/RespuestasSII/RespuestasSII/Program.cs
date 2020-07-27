using System;
using System.Collections.Generic;

using System.Configuration;
using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.IO;
using FirmaXML;
using System.Collections;
using System.Net.Mail;
using System.Xml.Serialization;
using System.Data.SqlClient;
using System.Data;

namespace RespuestasSII
{
    public class RespuestasSII
    {
        private static Hashtable configuracion = new Hashtable();
        public static DataTable tbEmisor;
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            /*  Application.EnableVisualStyles();
              Application.SetCompatibleTextRenderingDefault(false);
              Application.Run(new Form1());
             */
        }

        public string ramdonString(int largo)
        {
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            char[] stringChars = new char[10];
            var random = new Random();

            for (int i = 0; i < stringChars.Length; i++)
            {
                stringChars[i] = chars[random.Next(chars.Length)];
            }

            string s = new string(stringChars);
            return s;
        }
        public string[] Proceso(string CertPath, string CertPass, string RequiereProxy, string ProxyServer,
            string ProxyUser, string ProxyPass, string DtePath, string DtePathProc, string maximoIntentos,
            string DtePathProcNOK, string VidaSemillaMil, string[] DteEstOK, string OMITE_VAL_SII, string Valida_en_Sii, string Omitir_Guia, string Reintentos_Valida_Sii)

        {


              
           
            X509Certificate2 cert = new X509Certificate2(CertPath, CertPass);

            IWebProxy Proxya = System.Net.WebRequest.GetSystemWebProxy();
            if (RequiereProxy == "S")
            {
                NetworkCredential nc = new NetworkCredential(ProxyUser, ProxyPass, "");
                Proxya.Credentials = nc;
            }
            DirectoryInfo dir = new DirectoryInfo(DtePath);
           
            int numProcesados = 0;
            int numTotal = 0;
            Console.Clear();
            //semilla
            XmlDocument xmlSemilla = new XmlDocument();
            //
            Console.WriteLine("*****************************************************");
            Console.WriteLine("*****************************************************");
            Console.WriteLine("*******                                       *******");
            Console.WriteLine("*****       Actualizado 01-08-2019 (SII)        *****");
            Console.WriteLine("*******                                       *******");
            Console.WriteLine("*****************************************************");
            Console.WriteLine("*****************************************************");
            Console.WriteLine("CONSULTA DE ESTADO DE DTE");
            Console.WriteLine("PROCESO INICIADO (" + DateTime.Now.ToString() + ")...");
            int cant_a_procesar = dir.GetFiles("*.xml").Length;
            Console.WriteLine("XML DISPONIBLES A PROCESAR: " + cant_a_procesar.ToString());
            Console.WriteLine("*****************************************************");
            
            /******************************** HISTORIAL DE ACTUALIZACIONES **************************************
            01-08-2019 Se realiza actualizacion para omitir la verificacion de la guia de despacho en el SII


            ****************************** FIN HISTORIAL DE ACTUALIZACIONES ************************************/
            
            
            //
            //HORA SEMILLA
            DateTime horaSemilla = DateTime.Now;


            //Crea nuevo directorio si no existe 27-11-2018 (mgarrido)
            if(!Directory.Exists(Valida_en_Sii))
            {
                try{Directory.CreateDirectory(Valida_en_Sii);}
                catch { }
            }

            //Genera nuevos documentos  
            foreach (FileInfo f in dir.GetFiles("*.xml"))
            {
               
                string name = f.Name;
                string fullName = f.FullName;
                XmlDocument xmlDTE = new XmlDocument();
                xmlDTE.PreserveWhitespace = true;
                xmlDTE.Load(fullName);
                XmlNodeList elemDTE_Traspaso = xmlDTE.GetElementsByTagName("Documento");

                //separa xml de paquetes de envio.
                for (int ctDte = 0; ctDte <= elemDTE_Traspaso.Count - 1; ctDte++)
                {
                    try
                    {
                        string newNombre = Path.GetFileNameWithoutExtension(fullName);
                        try
                        {
                            if (newNombre.Length > 75) newNombre = newNombre.Substring(0, 74);
                        }
                        catch { }
                        if (!newNombre.Contains("_JDN2018_"))
                        {
                            string azar = "_JDN2018_" + ramdonString(8) + "_" + DateTime.Now.ToString("yyyyMMddThhmmss");
                            newNombre += azar;
                        }

                        newNombre += ".xml";

                        // Se gurdan en un nuevo directorio desde el 27-11-2018 (mgarrido)
                        XmlTextWriter aOuXML = new XmlTextWriter(Valida_en_Sii + "\\" + newNombre, System.Text.Encoding.GetEncoding("iso-8859-1"));
                        XmlDocument xmlEnc_traspaso = new XmlDocument();
                        xmlEnc_traspaso.PreserveWhitespace = true;
                        aOuXML.Formatting = Formatting.Indented;
                        aOuXML.WriteStartDocument();
                        aOuXML.WriteRaw(elemDTE_Traspaso[ctDte].OuterXml);
                        aOuXML.Flush();
                        aOuXML.Close();
                        System.Threading.Thread.Sleep(1000);
                    }
                    catch { }

                }
                try
                {
                    //mueve los dtes originales a esta carpeta
                    string dias = System.DateTime.Now.ToString("yyyyMMdd");
                    string carpetaORI = Path.GetDirectoryName(fullName) + "\\Historico_Originales\\" + dias + "\\";
                    DirectoryInfo dir1 = new DirectoryInfo(carpetaORI);
                    if (dir1.Exists == false)
                    {
                        dir1.Create();
                    }
                    File.Move(fullName, carpetaORI + Path.GetFileNameWithoutExtension(fullName) + DateTime.Now.Ticks.ToString() + ".xml");
                }
                catch
                {
                }
            }
            // Se lee desde el nuevo directorio desde el 27-11-2018 (mgarrido)
            DirectoryInfo NewDirectory = new DirectoryInfo(Valida_en_Sii);
            foreach (FileInfo f in NewDirectory.GetFiles("*.xml"))
            {
                int problemas = 0;

                #region LEO DTE
                string name = f.Name;
                string fullName = f.FullName;
                XmlDocument xmlDTE = new XmlDocument();
                xmlDTE.Load(fullName);

                XmlNodeList elemDTE = xmlDTE.GetElementsByTagName("Encabezado");

               
                //XmlDocument xmlEnc = new XmlDocument();
                //xmlEnc.InnerXml = elemDTE[0].OuterXml;
                //System.Threading.Thread.Sleep(2000);
                //xmlEnc.Save("Paso.xml");

                    
                    
                //    LectorXML lector = new LectorXML();
                //    Hashtable tbDTE = lector.lectorArchivo("Paso.xml");
                XmlDocument xmlEnc = new XmlDocument();
                xmlEnc.InnerXml = elemDTE[0].OuterXml;
                //System.Threading.Thread.Sleep(1500);
                //xmlEnc.Save("Paso.xml");

                MemoryStream xmlStream = new MemoryStream();
                xmlEnc.Save(xmlStream);

                xmlStream.Flush();//Adjust this if you want read your data 
                xmlStream.Position = 0;


                LectorXML lector = new LectorXML();
                //Hashtable tbDTE = lector.lectorArchivo("Paso.xml");
                Hashtable tbDTE = new Hashtable();
                try
                {
                    tbDTE = lector.lectorArchivo2(xmlStream);
                }
                catch
                {
                    System.Threading.Thread.Sleep(1500);
                    xmlEnc.Save("Paso.xml");
                    tbDTE = lector.lectorArchivo("Paso.xml");
                }


                    //obtengo datos de consulta
                    string[] arrayRutCons = tbDTE["RUTEmisor"].ToString().Split('-');
                    string RutConsultante = arrayRutCons[0];
                    string DvConsultante = arrayRutCons[1];
                    string[] arrayRutComp = tbDTE["RUTEmisor"].ToString().Split('-');
                    string RutCompania = arrayRutComp[0];
                    string DvCompania = arrayRutComp[1];
                    string[] arrayRutRece = tbDTE["RUTRecep"].ToString().Split('-');
                    string RutReceptor = arrayRutRece[0];
                    string DvReceptor = arrayRutRece[1];
                    string TipoDte = tbDTE["TipoDTE"].ToString();
                    string FolioDte = tbDTE["Folio"].ToString();
                    DateTime fecEmi = Convert.ToDateTime(tbDTE["FchEmis"]);
                    string FechaEmisionDte = fecEmi.ToString("ddMMyyyy");
                    string MontoDte = tbDTE["MntTotal"].ToString();

                    #region SEMILLA

                    string Si_Omite_Guia = "N";

                    if (Omitir_Guia == "S" && TipoDte == "52") {
                        Si_Omite_Guia = Omitir_Guia;
                    }
                    //semilla
                    string state = string.Empty;
                    int intentos = 0;
                    int resultMin = 0;
                    int maxIntentos = int.Parse(maximoIntentos);
                    while (intentos <= maxIntentos)
                    {
                        Console.WriteLine("Procesando XML : " + (numTotal + 1).ToString());
                        string estadoSemilla = string.Empty;
                        try
                        {
                            TimeSpan ts = DateTime.Now.Subtract(horaSemilla);
                            resultMin = ts.Milliseconds;

                            if ((resultMin == int.Parse(VidaSemillaMil)) || (resultMin == 0) || (xmlSemilla.InnerText.Length == 0))
                            {
                                cl.sii.semilla.CrSeedService mySemilla = new cl.sii.semilla.CrSeedService();
                                Console.WriteLine("Consultando SII-SEMILLA..." + DateTime.Now.ToString());
                                mySemilla.Timeout = 800;
                                if (RequiereProxy == "S") { mySemilla.Proxy = Proxya; }
                                // state = mySemilla.getState();
                                //obtengo hora
                                horaSemilla = DateTime.Now;
                                resultMin = 0;
                                //
                                xmlSemilla.InnerXml = mySemilla.getSeed();
                                intentos++;
                                Console.WriteLine("OK-SEMILLA..." + DateTime.Now.ToString());
                                break;
                            }
                            else
                            {
                                break;
                            }
                        }
                        catch (Exception e2)
                        {

                            intentos++;
                        }
                    }
                    if (xmlSemilla.InnerXml.Trim().Length > 0)
                    {
                        XmlNodeList elemList = xmlSemilla.GetElementsByTagName("SEMILLA");
                        string semilla = elemList[0].InnerText;
                        XmlDocument reqTokenSinFirma = new XmlDocument();
                        reqTokenSinFirma.InnerXml = "<?xml version='1.0'?><getToken><item><Semilla>" + semilla + "</Semilla></item></getToken>";

                        XmlDocument reqTokenConFirma = new XmlDocument();
                        FirmaXML.Firmado firmador = new Firmado();
                        reqTokenConFirma.InnerXml = firmador.Genera(reqTokenSinFirma, "", cert);
                    #endregion
                        #region TOKEN
                        //token
                        XmlDocument xmlToken = new XmlDocument();
                        cl.sii.token.GetTokenFromSeedService myToken = new cl.sii.token.GetTokenFromSeedService();
                        Console.WriteLine("Consultando SII-TOKEN...");
                        if (RequiereProxy == "S") { myToken.Proxy = Proxya; }
                        intentos = 0;
						if (Si_Omite_Guia == "N")
                        {
							while (intentos <= maxIntentos)
							{
                                try
                                {
                                    myToken.Timeout = 1000;
                                    xmlToken.InnerXml = myToken.getToken(reqTokenConFirma.InnerXml);
                                    XmlNodeList elemList2 = xmlToken.GetElementsByTagName("ESTADO");
                                    string estadoToken = elemList2[0].InnerText;
                                    if (estadoToken != "00")
                                    {
                                        xmlToken.InnerText = "";
                                        if (intentos <= maxIntentos) { state = string.Empty; }
                                    }
                                    else
                                    {
                                        Console.WriteLine("OK-TOKEN..." + DateTime.Now.ToString());
                                        break;
                                    }
                                    intentos++;

                                }
                                catch (Exception e1)
                                {
                                    intentos++;
                                }
                            }
						}
                        else
                        {
                            xmlToken.InnerXml = "<TOKEN>TOKEN OMITIDO</TOKEN>";
                            Console.WriteLine("VALIDACIÓN SII DTE OMITIDO (GUIA)");
                        }
                        if (xmlToken.InnerXml.Trim().Length > 0)
                        {
                            string token = string.Empty;
                            try
                            {
                                Console.WriteLine(xmlToken.InnerXml);
                                XmlNodeList elemList3 = xmlToken.GetElementsByTagName("TOKEN");
                                token = elemList3[0].InnerText;
                            }
                            catch(Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                            #endregion
                            #region RESPUESTA
                            //respuesta.
                            XmlDocument respuestaSII = new XmlDocument();

                            intentos = 0;
                            if (Si_Omite_Guia == "N")
                            {
                                while (intentos <= maxIntentos)
                                {
                                    try
                                    {
                                        Console.WriteLine("Consultando SII-ESTADO...");
                                        cl.sii.queryEstDte.QueryEstDteService myqueryEstDte = new cl.sii.queryEstDte.QueryEstDteService();
                                        myqueryEstDte.Timeout = 2000;
                                        if (RequiereProxy == "S") { myqueryEstDte.Proxy = Proxya; }

                                        respuestaSII.InnerXml = myqueryEstDte.getEstDte(RutConsultante, DvConsultante, RutCompania, DvCompania, RutReceptor, DvReceptor, TipoDte, FolioDte, FechaEmisionDte, MontoDte, token);
                                        Console.WriteLine("OK-ESTADO..." + DateTime.Now.ToString());
                                        System.Threading.Thread.Sleep(1000);
                                        break;
                                    }
                                    catch
                                    {
                                        intentos++;
                                    }
                                }
                            }
                            else
                            {
                                respuestaSII.InnerXml = "<TOKEN>RESPUESTA OMITIDO</TOKEN>";
                                Console.WriteLine("VALIDACIÓN SII DTE OMITIDO (GUIA)");
                            }
                            if (respuestaSII.InnerXml.Trim().Length > 0)
                            {
                                XmlNodeList elemList2;
                                string estadorespuestaSII = "";

                                if (Si_Omite_Guia == "N")
                                {
                                    try
                                    {

                                        //elemList2 = respuestaSII.GetElementsByTagName("GLOSA_ESTADO");
                                        elemList2 = respuestaSII.GetElementsByTagName("ESTADO");
                                        estadorespuestaSII = elemList2[0].InnerText;
                                    }
                                    catch
                                    {
                                        //elemList2 = respuestaSII.GetElementsByTagName("ESTADO");
                                        estadorespuestaSII = "0";
                                    }
                                }
                                if (OMITE_VAL_SII == "S")
                                {
                                    estadorespuestaSII = "DOK";
                                }
                                if (Si_Omite_Guia == "S")
                                {
                                    estadorespuestaSII = "DOK";
                                }

                                //valido si es una respuesta numerica. algo paso. Asi que lo dejo pendiente.
                                bool retornoNumero = false;
                                try
                                {
                                    double.Parse(estadorespuestaSII);
                                    retornoNumero = true;
                                }
                                catch
                                {
                                    retornoNumero = false;
                                }

                                if (!retornoNumero)
                                {
                                    int index = Array.IndexOf(DteEstOK, estadorespuestaSII);
                                    if (index > -1)  //estadorespuestaSII == "DOK" || estadorespuestaSII == "DNK")
                                    {
                                        //muevo archivos a pendientes
                                        //string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
                                        string dirProcesado = DtePathProc;
                                        DirectoryInfo dir2 = new DirectoryInfo(dirProcesado);
                                        if (!dir2.Exists) { dir2.Create(); }
                                        if (!File.Exists(dirProcesado + @"\" + f.Name))
                                        {
                                              f.MoveTo(dirProcesado + @"\" + f.Name);
                                        }
                                        else
                                        {
                                            f.MoveTo(dirProcesado + @"\" + "COPIA_" + DateTime.Now.Ticks.ToString("yyyMMddThhmmss") + "_" + f.Name);
                                            //f.MoveTo(dirProcesado + "COPIA_" + Path.GetFileNameWithoutExtension(fullName) + "_" + ctDte + DateTime.Now.Ticks.ToString() + ".xml");

                                        }
                                    }
                                    else
                                    {

                                        /*//muevo archivos a rechazados
                                        string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
                                        string dirProcesado = DtePathProcNOK + @"\" + carpetaDias;
                                        DirectoryInfo dir2 = new DirectoryInfo(dirProcesado);
                                        if (!dir2.Exists) { dir2.Create(); }
                                        f.MoveTo(dirProcesado + @"\" + f.Name);
                                        //guardo respuesta cuando es rechazado
                                        string newResp = f.Name.Substring(0, f.Name.Length - 4) + "_RES_" + DateTime.Now.Ticks + f.Extension;
                                        respuestaSII.Save(dirProcesado + @"\" + newResp);*/

                                        string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
                                        string dirProcesado = DtePathProcNOK + @"\" + carpetaDias;
                                        DirectoryInfo dir2 = new DirectoryInfo(dirProcesado);
                                        if (!dir2.Exists) { dir2.Create(); }
                                        string newResp = f.Name.Substring(0, f.Name.Length - 4) + "_RES_" + DateTime.Now.Ticks + f.Extension;

                                        var fechaRecepcion = f.LastWriteTime;
                                        DateTime Hoy = DateTime.Now;
                                        TimeSpan ts = Hoy - fechaRecepcion;
                                        int differenceInDays = ts.Days;

                                        if (estadorespuestaSII.Contains("FAU") && differenceInDays <= int.Parse(Reintentos_Valida_Sii))
                                        {
                                            Console.WriteLine("DOCUMENTO NO EXISTENTE EN SII, HAN PASADO " + differenceInDays + " DIAS");
                                        }
                                        else
                                        {
                                            Console.WriteLine("DOCUMENTO SE MUEVE A CARPETA DE RECHAZADOS");
                                            //muevo archivos a rechazados
                                            f.MoveTo(dirProcesado + @"\" + f.Name);
                                        }
                                        //guardo respuesta cuando es rechazado
                                        respuestaSII.Save(dirProcesado + @"\" + newResp);

                                    }
                                    numProcesados++;
                                }
                                else
                                {
                                    problemas++;
                                    Console.WriteLine("RETORNO NO VALIDO, ESTADO NUMERO (RESPUESTA), ARCHIVO: " + name);
                                }

                            }
                            else
                            {
                                problemas++;
                                Console.WriteLine("CONEXION A SII NO FUE POSIBLE (RESPUESTA), ARCHIVO: " + name);
                            }

                        }
                        else
                        {
                            problemas++;
                            Console.WriteLine("CONEXION A SII NO FUE POSIBLE (TOKEN), ARCHIVO: " + name);
                        }
                    }
                    else
                    {
                        problemas++;
                        Console.WriteLine("CONEXION A SII NO FUE POSIBLE (SEMILLA), ARCHIVO: " + name);
                    }
                            #endregion

                    numTotal++;
                

                #endregion
             
            }
            string[] myRetorno = new string[2];
            myRetorno[0] = numTotal.ToString();
            myRetorno[1] = numProcesados.ToString();
            Console.WriteLine("*****************************************************");
            Console.WriteLine("PROCESO FINALIZADO (" + DateTime.Now.ToString() + ").");
            Console.WriteLine("TOTAL: "+ numTotal.ToString());
            Console.WriteLine("PROCESADOS: " + numProcesados.ToString());
            Console.WriteLine("*****************************************************");

   
            return myRetorno;
           
        }


        //proceso de sujeto iva dic 2018
        public bool rechazoComercial(string rut, string tipo, string folio, string fecha, string monto, string nombre,
            string rutaCerti, string passCerti, string conex, string rutaTmp)
        {
            

            bool rsp = false;

            try
            {
                //cargo datos del emisor
                datosEmisor(conex);
                //cargo configuracion
                configuracionUsuario(conex);
                string rutEmisor = tbEmisor.Rows[0]["rut"].ToString();
                string rutReceptor = rut;
                string rutaCertificado = rutaCerti;
                string claveCertificado = passCerti;
                //recupero idenvio asociado
                PrototipoEfactura.prototipoEstadoDTEAV estadoAV = new PrototipoEfactura.prototipoEstadoDTEAV();
                string certificado = "S";
                string ambiente = "1";
                if (certificado == "S")
                {
                    ambiente = "2";
                }
               
                string TRACKID = DateTime.Now.ToString("hhmmssMMdd");
                string estado = "2";
                string accion = "";
                string BodyRechazo = "";
                if (estado == "0") accion = "Aprobacion Comercial";
                if (estado == "1") accion = "Aprobacion Comercial con Discrepacia";
                if (estado == "2")
                {
                    accion = "Rechazo Comercial";
                    BodyRechazo = "Se rechaza vuestra factura en atención a vuestra calidad de “Contribuyente sujeto a retención” y a la nuestra de “Agente Retenedor del IVA”. <br><br>" +
                    "Lo anterior considerando las instrucciones del Servicios de Impuestos Internos sobre cambio de sujeto total del IVA según lo previsto en los artículos 2º, N°s 3°) y 4°), 3º y 10 de la Ley sobre Impuesto a las Ventas y Servicios.";
                }
                //verifco si el contribuyente existe tabla de emisores
                Contribuyente myEmisor = new Contribuyente();
                myEmisor = getContribuyente(rut,conex);
                if (myEmisor.Rut == null)
                {
                    return false;
                }
                else
                {
                    PrototipoEfactura.prototipoResultadoDTE myprototipoResultadoDTE = new PrototipoEfactura.prototipoResultadoDTE(rutaCertificado, claveCertificado, rutReceptor, rutEmisor);
                    bool res1 = myprototipoResultadoDTE.addDTE(rut, tbEmisor.Rows[0]["rut"].ToString(),
                        tipo, folio, fecha, monto, estado, configuracion["RET_IVA_ASUNTO_GLOSA_RECHA_NORMAL"].ToString(), TRACKID, "");

                    string nomFile = accion.Replace(" ", "") + "R" + rut + "_T" + tipo + "_F" + folio + "_" + DateTime.Now.ToString("yyyyMMddThhmmss") + ".XML";
                    string salidaAceptadoOK = preparaRuta(rutaTmp  + @"\" ) + nomFile;
                    string resultado = myprototipoResultadoDTE.generaRespuesta(ref salidaAceptadoOK);
                   

                    Contribuyente myReceptor = new Contribuyente();
                    myReceptor = getContribuyente(tbEmisor.Rows[0]["rut"].ToString(), conex );
                    string asunto = accion + " - " + "R" + rut + "_T" + tipo + "_F" + folio;
                    string cuerpo = "Adjuntamos " + accion + ".<br><br>Atte. " + tbEmisor.Rows[0]["Razon"].ToString();
                    if (estado == "2")
                        cuerpo = BodyRechazo + ".<br><br>Atte. " + tbEmisor.Rows[0]["Razon"].ToString();
                    bool rsnv = enviaMail(myEmisor.Correo, myReceptor.Correo, asunto, salidaAceptadoOK, cuerpo, conex);
                    if (!rsnv)
                    {
                        rsp = false;
                    } 
                    else
                    {

                        //actulizar 
                        string ssql = "update RET_IVA_DTE_GENERACION_PND set aprobComercial='S',  fecAprobComercial='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' ";
                        ssql += "where rutEmisor='" + rut + "' and tipo='" + tipo + "' and folio='" + folio + "' ";
                        int kk = ejecutaSqlComando(ssql,conex);
                        //enviar
                        rsp = true;
                    }
                }
            }
            catch (Exception g1)
            {
               
                return false;

            }
            return rsp;
        }

        public bool rechazoSII(string rut, string tipo, string folio, string fecha, string monto, string rutaCerti,
            string passCerti, string conex)
        {
            bool rsp = false;
            try
            {
                //cargo configuracion
                configuracionUsuario(conex);
                string rutac = configuracion["RUTA_CERTIFICADO"].ToString();
                string passc = configuracion["CLAVE_CERTIFICADO"].ToString();


                PrototipoEfactura.prototipoJRN_AceptacionReclamoDoc
                    bknWS = new PrototipoEfactura.prototipoJRN_AceptacionReclamoDoc("2", rutac, passc);

                string ru1 = rut.Split('-')[0];
                string dv1 = rut.Split('-')[1];
                string ti1 = tipo; string fo1 = folio;
                string rs14 = bknWS.IngresarAceptacionReclamoDoc(ru1, dv1, ti1, fo1, "RCD");

                try
                {
                    XmlDocument xxml = new XmlDocument();
                    xxml.LoadXml(rs14);

                    string codigo = xxml.GetElementsByTagName("codResp")[0].InnerText;
                    string glosaestado = xxml.GetElementsByTagName("descResp")[0].InnerText;
                    if (codigo == "27") glosaestado += "\nACEPTACION REALIZADA.";
                   
                    //actulizar 
                    string ssql = "update RET_IVA_DTE_GENERACION_PND set aprobSii='S',  fecAprobSII='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' ";
                    ssql += "where rutEmisor='" + rut + "' and tipo='" + tipo + "' and folio='" + folio + "' ";
                    int kk = ejecutaSqlComando(ssql, conex);
                    //enviar
                    rsp = true;
                }
                catch { }

            }
            catch (Exception g1)
            {
                return false;
                
            }



            return rsp;
        }

        public string grabaPendienteRetener(string rut, string tipo, string folio, string fecha, string nombre, string neto, string iva,
                                         string total, XmlDocument xml, string conexionSql)
        {
            string rsp = "";
            string ticket = ya_Existe_Retencion(rut, tipo, folio, conexionSql);
            rsp = ticket;
             if (ticket.Trim().Length == 0)
             {
                 ticket = "RE" + rut.Replace("-","") + "T" + tipo + "F" + folio + "F" + DateTime.Now.ToString("yyyyMMddThhmmss");
                 string ssql = "INSERT INTO [RET_IVA_DTE_GENERACION_PND] ([estado],[rutEmisor],[tipo],[folio],[fecemi],[nombre],[neto],[iva],[total],[xml],[ticket]) VALUES (";
                 ssql += "'PND','" + rut + "','" + tipo + "','" + folio + "','" + fecha + "','" + nombre + "','" + neto + "','" + iva + "','" + total + "',";
                 ssql += "'" + xml.OuterXml + "','" + ticket + "' )";
                 int r = ejecutaSqlComando(ssql, conexionSql);
                 if (r == 1) rsp = ticket;
             }
            
            return rsp;
        }

        public bool enviaAlertaRetencion(string rut, string tipo, string folio, string fecha, string monto, XmlDocument  xml,
          string conex, string rutaTmp, string ticket, string nombre)
            {
            bool rsp = false;
            try
            {
                //cargo configuracion
                configuracionUsuario(conex);
                string ruta =preparaRuta(rutaTmp) + "Factura_E"+rut+"_T"+tipo +"_F"+folio + "_"+ DateTime.Now.ToString("yyyyMMddThhmmss")+".xml";
                xml.Save(ruta);
              
                string cuerpo ="Estimados:<br><br>" +
                               "El siguiente DTE ha sido rechazado por calificar como <b>SUJETO DE RETENCION DE IVA</b>:  " +
                               "<li>Rut: " +rut + "-" + nombre + "</li>" +
                               "<li>Tipo: " +tipo + "</li>" +
                               "<li>Folio: " +folio + "</li>" +
                               "<li>F.Emision: " + fecha + "</li>" +
                               "<li>Total: " +monto  + "</li><br>" +
                               "Favor ingrese al siguiente LINK, ingrese sus credenciales y a continuacion copie y pegue el siguiente <b>TICKET</b> " +
                               "en opcion de <b>Factura de Compra</b>:<br><br>" +
                               "<li>TICKET: <b>" + ticket + "</b>" +
                               "<li>LINK: <b>" + configuracion["RET_IVA_LINK_GENERACION"].ToString ()  + "</b>" +
                               "<br><br>Saludos" +
                               "<br><br>**NO CONTESTE ESTE MAIL, HA SIDO ENVIADO DE FORMA AUTOMATICA.";
                                
                
                bool env = enviaMail(configuracion["RET_IVA_ASUNTO_CORREOS_ALERTA"].ToString(),configuracion["RET_IVA_ASUNTO_CORREO_DESDE"].ToString(),
                                    configuracion["RET_IVA_ASUNTO_MENSAJE_RECHA_NORMAL"].ToString().Replace("[RUT]",rut ).Replace ("[FOLIO]", folio ),
                                    ruta  , cuerpo, conex  );

                rsp= env;

            }
            catch (Exception g1)
            {
                return false;

            }



            return rsp;
        }


        private static string preparaRuta(string _ruta)
        {
            string rsp = "";
            try
            {

                if (_ruta.Trim().Substring(_ruta.Length - 1, 1) != @"\")
                {
                    rsp = _ruta + @"\";
                }
                else
                {
                    rsp = _ruta;
                }

                if (!Directory.Exists(rsp))
                {
                    Directory.CreateDirectory(rsp);
                }

            }
            catch
            {
                rsp = "";
            }

            return rsp;
        }

        public int ejecutaSqlComando(string ssql, string conexionSql)
        {
            int rsp = 0;
            string sql1 = ssql;
            using (SqlConnection conn1 = new SqlConnection(conexionSql))
            {
                conn1.Open();
                SqlCommand comando1 = new SqlCommand(sql1, conn1);
                int count1 = Convert.ToInt32(comando1.ExecuteNonQuery());
                if (count1 != 1)
                {
                    rsp = count1;
                }
                else
                {
                    rsp = count1;
                }
            }

            return rsp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="para"></param>
        /// <param name="desde"></param>
        /// <param name="asunto"></param>
        /// <param name="fAdjunto">mas de uno separado por ;</param>
        /// <param name="cuerpo"></param>
        /// <returns></returns>
        private static bool enviaMail(string para, string desde, string asunto, string fAdjunto, string cuerpo, string conex)
        {

            configuracionUsuario(conex);
            bool rsp = true;
            try
            {
                using (MailMessage objMail = new MailMessage())
                {
                    objMail.From = new MailAddress(configuracion["RET_IVA_ASUNTO_CORREO_DESDE"].ToString()); //Remitente
                    if (configuracion["MAIL_CORREOS_DEPURACION"].ToString().Trim().Length > 0)
                    {
                        objMail.To.Add(configuracion["MAIL_CORREOS_DEPURACION"].ToString()); //Email a enviar 
                    }
                    else
                    {
                        objMail.To.Add(para); //Email a enviar 
                    }
                    //objMail.CC.Add("hfigueroa@jordan.cl"); //Email a enviar copia

                    if (configuracion["MAIL_CORREOS_COPIA_OCULTA"].ToString().Trim().Length > 0)
                    {
                        objMail.Bcc.Add(configuracion["MAIL_CORREOS_COPIA_OCULTA"].ToString()); //Email a enviar 
                    }
                    //else
                    //{
                    //    objMail.Bcc.Add(para); //Email a enviar 
                    //}
                    // objMail.Bcc.Add("hfigueroa@jordan.cl"); //Email a enviar oculto
                    objMail.Subject = asunto;
                    objMail.IsBodyHtml = true; //Formato Html del email
                    string[] adjuntos = fAdjunto.Split(';');
                    for (int k = 0; k < adjuntos.Length; k++)
                    {
                        if (adjuntos[k].Trim().Length > 0)
                        {
                            objMail.Attachments.Add(new Attachment(adjuntos[k]));
                        }
                    }
                    objMail.Body = cuerpo;
                    SmtpClient SmtpMail = new SmtpClient(configuracion["DTE_HOST_SMTP"].ToString(), int.Parse(configuracion["DTE_PORT_SMTP"].ToString()));
                    if (configuracion["DTE_USER_SMTP"].ToString().Trim().Length > 0)
                    {
                        SmtpMail.UseDefaultCredentials = false;
                        SmtpMail.EnableSsl = true;
                        SmtpMail.Credentials = new System.Net.NetworkCredential(configuracion["DTE_USER_SMTP"].ToString(), configuracion["DTE_PASSWORD_SMTP"].ToString());
                    }
                    SmtpMail.Send(objMail);
                    //SmtpMail..Dispose();
                }
            }
            catch
            {
                rsp = false;
            }


            return rsp;
        }

        public static void configuracionUsuario(string conex)
        {
            configuracion = new Hashtable();
            string con = conex;
            SqlConnection sqlCnx = new SqlConnection();
            try
            {
                configuracion.Clear();
                SqlCommand sqlCmd = new SqlCommand();
                sqlCnx.ConnectionString = con;
                sqlCmd.Connection = sqlCnx;
                sqlCnx.Open();
                sqlCmd.CommandTimeout = 999999999;
                sqlCmd.CommandText = "select * from configuracion ";
                DataTable _tabla = new DataTable();
                _tabla.TableName = "tabla";
                _tabla.Load(sqlCmd.ExecuteReader());
                foreach (DataRow sqlDrTC in _tabla.Rows)
                {
                    if (!configuracion.ContainsKey(sqlDrTC["key"].ToString()))
                    {
                        configuracion.Add(sqlDrTC["key"].ToString(), sqlDrTC["value"].ToString());
                    }

                }
            }
            catch (Exception ex)
            {
                //          throw ex;
            }
            finally
            {
                sqlCnx.Close();
            }
        }

        private static Contribuyente getContribuyente(string rut, string conex)
        {
            Contribuyente myContribuyente = new Contribuyente();
            string con = conex;
            SqlConnection sqlCnx = new SqlConnection();
            DataTable _tabla;
            try
            {
                SqlCommand sqlCmd = new SqlCommand();
                sqlCnx.ConnectionString = con;
                sqlCmd.Connection = sqlCnx;
                sqlCmd.CommandText = "select * from ReceptoresElectronicos where rut='" + rut.Trim() + "'";
                sqlCnx.Open();
                sqlCmd.CommandTimeout = 999999999;
                _tabla = new DataTable();
                _tabla.Load(sqlCmd.ExecuteReader());
                int cantReg = _tabla.Rows.Count;
                foreach (DataRow sqlDrTC in _tabla.Rows)
                {
                    myContribuyente.Rut = sqlDrTC["rut"].ToString();
                    myContribuyente.Nombre = sqlDrTC["razon_social"].ToString();
                    myContribuyente.Correo = sqlDrTC["email"].ToString();
                }

            }
            catch (Exception ex)
            {
                myContribuyente = null;
            }
            finally
            {
                sqlCnx.Close();
            }
            return myContribuyente;
        }

        public static void datosEmisor(string conex)
        {
            string con = conex;
            SqlConnection sqlCnx = new SqlConnection();
            try
            {
                SqlCommand sqlCmd = new SqlCommand();
                sqlCnx.ConnectionString = con;
                sqlCmd.Connection = sqlCnx;
                sqlCnx.Open();
                //traigo malla de indicador por evaluacion
                //   sqlCnx.Open();
                sqlCmd.CommandTimeout = 999999999;
                try
                {
                    sqlCmd.CommandText = "select top 1 * from emisor where Matriz='S'";
                    tbEmisor = new DataTable();
                    tbEmisor.TableName = "tabla";
                    tbEmisor.Load(sqlCmd.ExecuteReader());
                }
                catch
                {
                    sqlCmd.CommandText = "select top 1 * from emisor ";
                    tbEmisor = new DataTable();
                    tbEmisor.TableName = "tabla";
                    tbEmisor.Load(sqlCmd.ExecuteReader());
                }
            }
            catch
            {

            }
            finally
            {
                sqlCnx.Close();
            }
        }

        public static string ya_Existe_Retencion(string rut, string tipo, string folio, string conex)
        {
            string resp = "";
            string con = conex;
            SqlConnection sqlCnx = new SqlConnection();
            try
            {
                SqlCommand sqlCmd = new SqlCommand();
                sqlCnx.ConnectionString = con;
                sqlCmd.Connection = sqlCnx;
                sqlCnx.Open();
                //traigo malla de indicador por evaluacion
                //   sqlCnx.Open();
                sqlCmd.CommandTimeout = 999999999;
                try
                {
                    sqlCmd.CommandText = "select ticket from RET_IVA_DTE_GENERACION_PND where rutEmisor='" + rut + "' and tipo ='" + tipo + "' and folio='" + folio + "' ";
                    DataTable tb = new DataTable();
                    tb.TableName = "tabla";
                    tb.Load(sqlCmd.ExecuteReader());
                    if (tb.Rows.Count > 0)
                    {
                        resp = tb.Rows[0]["ticket"].ToString();
                    }
                }
                catch
                {
                }
            }
            catch
            {

            }
            finally
            {
                sqlCnx.Close();
            }

            return resp;
        }


        public class Contribuyente
        {
            public string Rut;
            public string Nombre;
            public string Correo;
        }

    }



   


    
    

    }


