using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;


    public class ResultadoDTE
    {

       string _TipoDTE, _Folio, _FchEmis, _RUTEmisor, _RUTRecep, _MntTotal, _EstadoDTE, _EstadoDTEGlosa, _RznSocial, _Codigo, _MailRznSocial, _CodEnvio;

         public void creaXML()
             //(string TipoDTE, string Folio, string FchEmis, string RUTEmisor, string RUTRecep,
         //     string MntTotal, string EstadoDTE, string EstadoDTEGlosa, string RznSocial, string Codigo, string MailRznSocial)
         {


         }

        public string TipoDTE
        {
            get { return _TipoDTE; }
            set { _TipoDTE = value; }
        }


        public string Folio
        {
            get { return _Folio; }
            set { _Folio = value; }
        }

        public string FchEmis
        {
            get { return _FchEmis; }
            set { _FchEmis = value; }
        }

        public string RUTEmisor
        {
            get { return _RUTEmisor; }
            set { _RUTEmisor = value; }
        }
  
        public string RUTRecep
        {
            get { return _RUTRecep; }
            set { _RUTRecep = value; }
        }
        public string MntTotal
        {
            get { return _MntTotal; }
            set { _MntTotal = value; }
        }
   
       public string CodEnvio
        {
            get { return _CodEnvio; }
            set { _CodEnvio = value; }
        }

       

        public string EstadoDTE
        {
            get { return _EstadoDTE; }
            set { _EstadoDTE = value; }
        }
        public string RznSocial
        {
            get { return _RznSocial; }
            set { _RznSocial = value; }
        }
        public string EstadoDTEGlosa
        {
            get { return _EstadoDTEGlosa; }
            set { _EstadoDTEGlosa = value; }
        }
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        public string MailRznSocial
        {
            get { return _MailRznSocial; }
            set { _MailRznSocial = value; }
        }

    }

    public partial class Caratula
    {
        
        string _RutResponde, _RutRecibe, _IdRespuesta, _NroDetalles, _TmstFirmaResp;
        private decimal versionField;

        public  Caratula()
        {
            this.versionField = ((decimal)(1.0m));

        }

        public string RutResponde
        {
            get { return _RutResponde; }
            set { _RutResponde = value; }
        }


        public string RutRecibe
        {
            get { return _RutRecibe; }
            set { _RutRecibe = value; }
        }

        public string IdRespuesta
        {
            get { return _IdRespuesta; }
            set { _IdRespuesta = value; }
        }

        public string NroDetalles
        {
            get { return _NroDetalles; }
            set { _NroDetalles = value; }
        }

        public string TmstFirmaResp
        {
            get { return _TmstFirmaResp; }
            set { _TmstFirmaResp = value; }
        }
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal version
        {
            get
            {
                return this.versionField;
            }
            set
            {
                this.versionField = value;
            }
        }
    }

//    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.42")]
  //  [System.SerializableAttribute()]
//    [System.Diagnostics.DebuggerStepThroughAttribute()]
 //   [System.ComponentModel.DesignerCategoryAttribute("code")]
 //[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://www.sii.cl/SiiDte")]
  //  [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://www.sii.cl/SiiDte",IsNullable = true)]
    public class Resultado
    {
        Caratula _caratula;
        ResultadoDTE _resultadoDTE;
        private string _id;
        public Resultado()
        {
        }
        public Caratula Caratula
        {
            get { return _caratula; }
            set { _caratula = value; }
        }
        public ResultadoDTE ResultadoDTE
        {
            get { return _resultadoDTE; }
            set { _resultadoDTE = value; }
        }
        [System.Xml.Serialization.XmlAttributeAttribute(DataType = "ID")]
        public string ID
        {
            get { return _id; }
            set { _id = value; }
        }

    }

    public class RespuestaDTE
    {
        Resultado _resultado;
        public RespuestaDTE()
        {
        }
        public Resultado Resultado
        {
            get { return _resultado; }
            set { _resultado = value; }
        }
    }
