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


namespace RespuestasSII
{
    public class RespuestasSII
    {
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

        public string[] Proceso(string CertPath, string CertPass, string RequiereProxy, string ProxyServer,
            string ProxyUser, string ProxyPass, string DtePath, string DtePathProc, string maximoIntentos,
            string DtePathProcNOK, string VidaSemillaMil, string DteEstOK)

        {
            /*
           string CertPath = Properties.Settings.Default["CertPath"].ToString();
            string CertPass = Properties.Settings.Default["CertPass"].ToString();
            string RequiereProxy = Properties.Settings.Default["RequiereProxy"].ToString();
            string ProxyServer = Properties.Settings.Default["ProxyServer"].ToString();
            string ProxyUser = Properties.Settings.Default["ProxyUser"].ToString();
            string ProxyPass = Properties.Settings.Default["ProxyPass"].ToString();
            string DtePath = Properties.Settings.Default["DtePath"].ToString();
             * */
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
            Console.WriteLine("CONSULTA DE ESTADO DE DTE");
            Console.WriteLine("PROCESO INICIADO (" + DateTime.Now.ToString() + ")...");
            Console.WriteLine("*****************************************************");
            //
            //HORA SEMILLA
            DateTime horaSemilla = DateTime.Now;

            //Genera nuevos documentos  
            foreach (FileInfo f in dir.GetFiles("*.xml"))
            {
               
                string name = f.Name;
                string fullName = f.FullName;
                XmlDocument xmlDTE = new XmlDocument();
                xmlDTE.Load(fullName);
                XmlNodeList elemDTE_Traspaso = xmlDTE.GetElementsByTagName("Documento");

                for (int ctDte = 0; ctDte <= elemDTE_Traspaso.Count - 1; ctDte++)
                {
                    XmlTextWriter aOuXML = new XmlTextWriter(Path.GetDirectoryName(fullName) + "\\" + Path.GetFileNameWithoutExtension(fullName) + "_" + ctDte + ".xml", System.Text.Encoding.GetEncoding("iso-8859-1"));
                    XmlDocument xmlEnc_traspaso = new XmlDocument();
                    
                    xmlEnc_traspaso.PreserveWhitespace = true;
                    aOuXML.Formatting = Formatting.Indented;
                    aOuXML.WriteStartDocument();
                    aOuXML.WriteRaw(elemDTE_Traspaso[ctDte].OuterXml);
                    aOuXML.Flush();
                    aOuXML.Close();

                }
                //mueve los dtes originales a esta carpeta
                string dias = System.DateTime.Now.ToString("yyyyMMdd");
                string carpetaORI =  Path.GetDirectoryName(fullName) + "\\Historico_Originales\\" + dias + "\\";
                DirectoryInfo dir1 = new DirectoryInfo(carpetaORI);
                if (dir1.Exists == false) 
                {
                    dir1.Create();
                }
                File.Move(fullName, carpetaORI + Path.GetFileNameWithoutExtension(fullName) + DateTime.Now.Ticks.ToString() + ".xml");
            }
            foreach (FileInfo f in dir.GetFiles("*.xml"))
            {
                int problemas = 0;

                #region LEO DTE
                string name = f.Name;
                string fullName = f.FullName;
                XmlDocument xmlDTE = new XmlDocument();
                xmlDTE.Load(fullName);

                XmlNodeList elemDTE = xmlDTE.GetElementsByTagName("Encabezado");

               
                    XmlDocument xmlEnc = new XmlDocument();
                    xmlEnc.InnerXml = elemDTE[0].OuterXml;
                    xmlEnc.Save("Paso.xml");

                    
                    
                    LectorXML lector = new LectorXML();
                    Hashtable tbDTE = lector.lectorArchivo("Paso.xml");

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

                    //semilla
                    string state = string.Empty;
                    int intentos = 0;
                    int resultMin = 0;
                    int maxIntentos = int.Parse(maximoIntentos);
                    while (intentos <= maxIntentos)
                    {
                        string estadoSemilla = string.Empty;
                        try
                        {
                            TimeSpan ts = DateTime.Now.Subtract(horaSemilla);
                            resultMin = ts.Milliseconds;

                            if ((resultMin == int.Parse(VidaSemillaMil)) || (resultMin == 0) || (xmlSemilla.InnerText.Length == 0))
                            {
                                cl.sii.semilla.CrSeedService mySemilla = new cl.sii.semilla.CrSeedService();
                                Console.WriteLine("Consultando SII-SEMILLA..." + DateTime.Now.ToString());
                                mySemilla.Timeout = 1000;
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
                        if (xmlToken.InnerXml.Trim().Length > 0)
                        {

                            XmlNodeList elemList3 = xmlToken.GetElementsByTagName("TOKEN");
                            string token = elemList3[0].InnerText;

                        #endregion
                            #region RESPUESTA
                            //respuesta.
                            XmlDocument respuestaSII = new XmlDocument();

                            intentos = 0;
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
                                    break;
                                }
                                catch
                                {
                                    intentos++;
                                }
                            }
                            if (respuestaSII.InnerXml.Trim().Length > 0)
                            {
                                XmlNodeList elemList2;
                                string estadorespuestaSII = "";
                                try
                                {
                                    elemList2 = respuestaSII.GetElementsByTagName("GLOSA_ESTADO");
                                    estadorespuestaSII = elemList2[0].InnerText;
                                }
                                catch
                                {
                                    elemList2 = respuestaSII.GetElementsByTagName("ESTADO");
                                    estadorespuestaSII = elemList2[0].InnerText;
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
                                    if (DteEstOK.Contains(estadorespuestaSII))  //estadorespuestaSII == "DOK" || estadorespuestaSII == "DNK")
                                    {
                                        //muevo archivos a procesados
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
                                            f.MoveTo(dirProcesado + @"\" + "COPIA_" + DateTime.Now.Ticks.ToString() + "_" + f.Name);
                                            //f.MoveTo(dirProcesado + "COPIA_" + Path.GetFileNameWithoutExtension(fullName) + "_" + ctDte + DateTime.Now.Ticks.ToString() + ".xml");

                                        }
                                    }
                                    else
                                    {

                                        //muevo archivos a procesados
                                        string carpetaDias = DateTime.Now.ToString("yyyyMMdd");
                                        string dirProcesado = DtePathProcNOK + @"\" + carpetaDias;
                                        DirectoryInfo dir2 = new DirectoryInfo(dirProcesado);
                                        if (!dir2.Exists) { dir2.Create(); }
                                        f.MoveTo(dirProcesado + @"\" + f.Name);
                                        //guardo respuesta cuando es rechazado
                                        string newResp = f.Name.Substring(0, f.Name.Length - 4) + "_RES_" + DateTime.Now.Ticks + f.Extension;
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
            Console.WriteLine("*****************************************************");

   
            return myRetorno;
           
        }
    }

    }


