    using Microsoft.VisualBasic;
    using System.Xml;
    using System.Collections;
    using System.IO;
    using System.Text;
    public class LectorXML
    {
        private Hashtable tabla = new Hashtable();

        public Hashtable lectorString(string texto)
        {
            XmlTextReader reader = null;
            
            reader = new XmlTextReader(new MemoryStream(ASCIIEncoding.Default.GetBytes(texto)));
            // es un string
            string atributo;
            string nombre;
            atributo = string.Empty;
            nombre = string.Empty;
            while (reader.Read())
            {
                if ((reader.NodeType == XmlNodeType.Element))
                {
                    nombre = reader.Name;
 
                }
                else if ((reader.NodeType == XmlNodeType.Text))
                {
                    atributo = reader.Value.Replace(("" + ("\r\n" + "")), string.Empty).Trim();
                }
                if (((atributo != string.Empty)
                            && (nombre != string.Empty)))
                {
                    tabla.Add(nombre, atributo);
                    atributo = string.Empty;
                    nombre = string.Empty;
                }
            }
            reader.Close();
            return tabla;
        }

        public Hashtable lectorArchivo(string archivo)
        {
            XmlTextReader reader = null;
            reader = new XmlTextReader(new StreamReader(archivo, Encoding.GetEncoding("ISO-8859-9")));
             
            // es un path    Dim atributo, nombre As String
            string atributo;
            string nombre;
            atributo = string.Empty;
            nombre = string.Empty;
            while (reader.Read())
            {
                if ((reader.NodeType == XmlNodeType.Element))
                {
                    nombre = reader.Name;
                }
                else if ((reader.NodeType == XmlNodeType.Text))
                {
                    atributo = reader.Value.Replace(("" + ("\r\n" + "")), string.Empty).Trim();
                }
                if (((atributo != string.Empty)
                            && (nombre != string.Empty)))
                {
                    if (!tabla.ContainsKey(nombre)) {
                    tabla.Add(nombre, atributo);
                    }
                    atributo = string.Empty;
                    nombre = string.Empty;
                }
            }
            reader.Close();
            return tabla;
        }
    }







