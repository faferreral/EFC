using OpenHtmlToPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using GeneXus.Utils;
/// <summary>
/// Descripción breve de type_SdtTableToJson
/// </summary>
namespace GeneXus.Programs
{
    public class type_SdtTableToJson
    {
		private GxSimpleCollection<string> AV18to ;
        public type_SdtTableToJson()
        {
			string r = "";
        }
		public string GetTablesFromHTML(string html)
        {
            List<string> tables = new List<string>();
            html = html.Replace("<br>", "");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.OptionWriteEmptyNodes = true;
            doc.OptionAutoCloseOnEnd = true;
            doc.OptionOutputAsXml = true;
            doc.LoadHtml(html);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;

            foreach (HtmlAgilityPack.HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
            {
                tables.Add(FormatHtml(table.OuterHtml));
            }
            return Newtonsoft.Json.JsonConvert.SerializeObject(tables);
        }
        public String GetHeaderFromTable(string html)
        {
			html = System.Net.WebUtility.HtmlDecode(html);
            List<TableHeader> header = new List<TableHeader>();
            html = html.Replace("<br>", "");
            html = html.Replace("\\n", " ");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.OptionWriteEmptyNodes = true;
            doc.OptionAutoCloseOnEnd = true;
            doc.OptionOutputAsXml = true;
            doc.LoadHtml(html);
            //System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            var tableRows = doc.DocumentNode.SelectNodes("//tbody/tr");
            var columns = tableRows[0].SelectNodes("td|th");

            for (int e = 0; e < columns.Count; e++)
            {
                var value = columns[e].InnerText.Trim();
                TableHeader h = new TableHeader();
                h.Header = value;
                if (value != null)
                    header.Add(h);
            }
            return Newtonsoft.Json.JsonConvert.SerializeObject(header);
        }
		public string GetValueXPath(string html, string xpath)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            var q = doc.DocumentNode.SelectNodes(xpath);
            if (q != null)
                if (q.Count > 0)
                    return q.First().InnerHtml;
                else
                    return "";
            else
                return "";
        }
        public String ConvertTableToJson(string html, string xsd, string EmailProd)
        {
			html = System.Net.WebUtility.HtmlDecode(html);
			EmailProd = System.Net.WebUtility.HtmlDecode(EmailProd);
            html = html.Replace("<br>", "");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.OptionWriteEmptyNodes = true;
            doc.OptionAutoCloseOnEnd = true;
            doc.OptionOutputAsXml = true;
            doc.LoadHtml(html);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;

            foreach (HtmlAgilityPack.HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
            {
                //listatablas.Add(table.OuterHtml);
                bool r = ValidateSchema(table.OuterHtml, xsd);
                if (r)
                {
                    tabla_correcta = table.OuterHtml;
                    break;
                }
            }
            if (!String.IsNullOrEmpty(tabla_correcta))
            {
                HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                doc2.OptionFixNestedTags = true;
                doc2.OptionWriteEmptyNodes = true;
                doc2.OptionAutoCloseOnEnd = true;
                doc2.OptionOutputAsXml = true;
                doc2.LoadHtml(tabla_correcta);
                var tableRows = doc2.DocumentNode.SelectNodes("//tbody/tr");
                var columns = tableRows[0].SelectNodes("td|th");

                for (int e = 0; e < columns.Count; e++)
                {
                    var value = columns[e].InnerText.Trim();
                    if (value != null)
                        header.Add(value);
                }
                foreach (HtmlAgilityPack.HtmlNode table in doc2.DocumentNode.SelectNodes("//table"))
                {
                    foreach (HtmlAgilityPack.HtmlNode row in table.SelectNodes("tbody/tr|tr").Skip(1))
                    {
                        if (row.SelectNodes("th|td").Count == columns.Count)
                        {

                            List<String> values = new List<string>();
                            //listatablas.Add(row.SelectNodes("th|td").Count.ToString());
                            foreach (HtmlAgilityPack.HtmlNode cell in row.SelectNodes("th|td"))
                            {
                                values.Add(cell.InnerText.Trim());
                            }

                            var cotizacion = header.Zip(values, (nombre, valor) => String.Format("\"{0}\" : \"{1}\",", nombre.Trim(), valor.Trim()));
                            string elemento = string.Empty;
                            foreach (var item in cotizacion)
                            {
                                elemento += item;
                            }

                            contenido += String.Format(" {{ {0} }},", elemento);
                        }

                    }
                }

                json = String.Format("[{0}]", contenido);
            }
            log("EmailProd: " + EmailProd);
            return BuscarCampoJson(json, EmailProd);
        }
		
		public String ConvertTableToJsonXSLT(string html_original, string xsd_original, string xslt, string EmailProd, string patron)
        {
			html_original = System.Net.WebUtility.HtmlDecode(html_original);
			EmailProd = System.Net.WebUtility.HtmlDecode(EmailProd);
			EmailProd = System.Net.WebUtility.HtmlDecode(EmailProd);
            string tabla_correcta_original = GetTableHTML_XSD(html_original, xsd_original, xslt, patron);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;

            if (!String.IsNullOrEmpty(tabla_correcta_original))
            {
                //tabla_correcta = TransformDocument(tabla_correcta_original, xslt);
                tabla_correcta = tabla_correcta_original;
                HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                doc2.OptionFixNestedTags = true;
                doc2.OptionWriteEmptyNodes = true;
                doc2.OptionAutoCloseOnEnd = true;
                doc2.OptionOutputAsXml = true;
                doc2.LoadHtml(tabla_correcta);
                var tableRows = doc2.DocumentNode.SelectNodes("//tbody/tr");
                var columns = tableRows[0].SelectNodes("td|th");

                for (int e = 0; e < columns.Count; e++)
                {
                    var value = columns[e].InnerText.Trim();
                    if (value != null)
                        header.Add(value);
                }
                foreach (HtmlAgilityPack.HtmlNode table in doc2.DocumentNode.SelectNodes("//table"))
                {
                    foreach (HtmlAgilityPack.HtmlNode row in table.SelectNodes("tbody/tr|tr").Skip(1))
                    {
                        if (row.SelectNodes("th|td") != null)
                            if (row.SelectNodes("th|td").Count == columns.Count)
                            {

                                List<String> values = new List<string>();
                                //listatablas.Add(row.SelectNodes("th|td").Count.ToString());
                                foreach (HtmlAgilityPack.HtmlNode cell in row.SelectNodes("th|td"))
                                {
                                    values.Add(cell.InnerText.Trim());
                                }

                                var cotizacion = header.Zip(values, (nombre, valor) => String.Format("\"{0}\" : \"{1}\",", nombre.Trim(), valor.Trim().Replace("\"", "")));
                                string elemento = string.Empty;
                                foreach (var item in cotizacion)
                                {
                                    elemento += item;
                                }

                                contenido += String.Format(" {{ {0} }},", elemento);
                            }
                    }
                }

                json = String.Format("[{0}]", contenido);
            }
            return BuscarCampoJson(json, EmailProd);
        }
		public String ConvertTableToJson(string html)
        {
			html = System.Net.WebUtility.HtmlDecode(html);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;

            if (!String.IsNullOrEmpty(html))
            {
                HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                doc2.OptionFixNestedTags = true;
                doc2.OptionWriteEmptyNodes = true;
                doc2.OptionAutoCloseOnEnd = true;
                doc2.OptionOutputAsXml = true;
                doc2.LoadHtml(html);
                var tableRows = doc2.DocumentNode.SelectNodes("//tbody/tr");
                var columns = tableRows[0].SelectNodes("td|th");

                for (int e = 0; e < columns.Count; e++)
                {
                    var value = columns[e].InnerText.Trim();
                    if (value != null)
                        header.Add(value);
                }

                foreach (HtmlAgilityPack.HtmlNode table in doc2.DocumentNode.SelectNodes("//table"))
                {
                    foreach (HtmlAgilityPack.HtmlNode row in table.SelectNodes("tbody/tr|tr").Skip(1))
                    {
                        if (row.SelectNodes("th|td").Count == columns.Count)
                        {

                            List<String> values = new List<string>();
                            foreach (HtmlAgilityPack.HtmlNode cell in row.SelectNodes("th|td"))
                            {
                                values.Add(cell.InnerText.Trim().Replace("\"", ""));
                            }

                            var cotizacion = header.Zip(values, (nombre, valor) => String.Format("\"{0}\" : \"{1}\",", nombre.Trim(), valor.Trim().Replace( "\r\n", " " )
                  .Replace( "\r", " " )
                  .Replace( "\n", " " )));
                            string elemento = string.Empty;
                            foreach (var item in cotizacion)
                            {
								if (item != "\"\" : \"\",")
									elemento += item;
                            }

                            contenido += String.Format(" {{ {0} }},", elemento);
                        }

                    }
                }

                json = String.Format("[{0}]", contenido);

            }
            return json;
        }
		public String ConvertTableTheadToJson(string html)
        {
            html = System.Net.WebUtility.HtmlDecode(html);
            string xslt = "<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\"><xsl:output omit-xml-declaration=\"yes\" indent=\"yes\"/><xsl:template match=\"table\"><table border=\"1\" class=\"dataframe\"><tbody><xsl:for-each select=\"thead/tr|tbody/tr|tr\"><tr><xsl:for-each select=\"td|th\"><td><xsl:value-of select=\".\"/><xsl:value-of select=\"div/div/input/@value\"/></td></xsl:for-each></tr></xsl:for-each></tbody></table></xsl:template></xsl:stylesheet>";
            html = TransformDocument(html, xslt);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;

            if (!String.IsNullOrEmpty(html))
            {
                HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                doc2.OptionFixNestedTags = true;
                doc2.OptionWriteEmptyNodes = true;
                doc2.OptionAutoCloseOnEnd = true;
                doc2.OptionOutputAsXml = true;
                doc2.LoadHtml(html);
                var tableRows = doc2.DocumentNode.SelectNodes("//tbody/tr");
                var columns = tableRows[0].SelectNodes("td|th");

                for (int e = 0; e < columns.Count; e++)
                {
                    var value = columns[e].InnerText.Trim();
                    if (value != null && !string.IsNullOrEmpty(value))
                        header.Add(value);
                }

                foreach (HtmlAgilityPack.HtmlNode table in doc2.DocumentNode.SelectNodes("//table"))
                {
                    foreach (HtmlAgilityPack.HtmlNode row in table.SelectNodes("tbody/tr|tr").Skip(1))
                    {
                        if (row.SelectNodes("th|td").Count == columns.Count)
                        {
                            List<String> values = new List<string>();
                            foreach (HtmlAgilityPack.HtmlNode cell in row.SelectNodes("th|td"))
                            {
                                if (!string.IsNullOrEmpty(cell.InnerText))
                                    values.Add(cell.InnerText.Trim().Replace("\"", ""));
                                else
                                {
                                    var celda = cell;
                                }
                            }

                            var cotizacion = header.Zip(values, (nombre, valor) => String.Format("\"{0}\" : \"{1}\",", nombre.Trim(), valor.Trim()));
                            string elemento = string.Empty;
                            foreach (var item in cotizacion)
                            {
                                elemento += item;
                            }

                            contenido += String.Format(" {{ {0} }},", elemento);
                        }

                    }
                }

                json = String.Format("[{0}]", contenido);

            }
            return json;
        }
		public string TransformDocument(string doc, string stylesheet)
        {
            Func<string, XmlDocument> GetXmlDocument = (xmlContent) =>
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlContent);
                return xmlDocument;
            };

            try
            {
                var document = GetXmlDocument(doc);
                var style = GetXmlDocument(stylesheet);

                System.Xml.Xsl.XslCompiledTransform transform = new System.Xml.Xsl.XslCompiledTransform();
                transform.Load(style); // compiled stylesheet
                System.IO.StringWriter writer = new System.IO.StringWriter();
                XmlReader xmlReadB = new XmlTextReader(new StringReader(document.DocumentElement.OuterXml));
                transform.Transform(xmlReadB, null, writer);
                return writer.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
		public String GetTableHTML_XSD(string html, string xsd, string xslt, string patron)
        {
            html = html.Replace("<br>", "");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.OptionWriteEmptyNodes = true;
            doc.OptionAutoCloseOnEnd = true;
            doc.OptionOutputAsXml = true;
            doc.LoadHtml(html);
            System.Collections.Generic.List<String> header = new System.Collections.Generic.List<string>();
            string tabla_correcta = string.Empty;
            string json = string.Empty;
            string contenido = string.Empty;
			bool resultado = false;
 
            foreach (HtmlAgilityPack.HtmlNode table in doc.DocumentNode.SelectNodes("//table"))
            {
                tabla_correcta = string.Empty;
				log(patron);
                if (String.IsNullOrEmpty(patron))
                {
                    tabla_correcta = TransformDocument(table.OuterHtml, xslt);
					tabla_correcta = DeleteEmptyTrTag(tabla_correcta);
                    bool r = ValidateSchema(tabla_correcta, xsd);
					resultado = r;
                    if (r)
                    {
                        //tabla_correcta = table.OuterHtml;
                        break;
                    }
                }
                else
                {
                    
                    if (table.OuterHtml.Contains(patron))
                    {
                        HtmlAgilityPack.HtmlDocument doct = new HtmlAgilityPack.HtmlDocument();
                        doct.OptionFixNestedTags = true;
                        doct.OptionWriteEmptyNodes = true;
                        doct.OptionAutoCloseOnEnd = true;
                        doct.OptionOutputAsXml = true;
                        doct.LoadHtml(table.OuterHtml);

                        int tables = doct.DocumentNode.SelectNodes("//table").Count;
                        if (tables == 1)
                        {
                            tabla_correcta = TransformDocument(table.OuterHtml, xslt);
                            tabla_correcta = DeleteEmptyTag(tabla_correcta);
                            bool r = ValidateSchema(tabla_correcta, xsd);
							resultado = r;
                            if (r)
                            {
                                //tabla_correcta = table.OuterHtml;
                                break;
                            }
                        }
                        
                    }
                }
            }
			//log("GetTableHTML_XSD tabla_correcta: " + tabla_correcta);
			if (resultado == false)
                tabla_correcta = "";
            return tabla_correcta;
        }
        public bool ValidateSchema(string strxml, string xsd)
        {
            strxml = strxml.Replace("&nbsp;", "");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.OptionFixNestedTags = true;
            doc.LoadHtml(strxml);

            System.Xml.XmlReader xmlReaderHtml = System.Xml.XmlReader.Create(new System.IO.StringReader(doc.Text));
            System.Xml.XmlDocument xml = new System.Xml.XmlDocument();
            xml.Load(xmlReaderHtml);
            System.Xml.XmlReader xmlReader = System.Xml.XmlReader.Create(new System.IO.StringReader(xsd));
            xml.Schemas.Add(null, xmlReader);

            try
            {
                xml.Validate(null);
            }
            catch (System.Xml.Schema.XmlSchemaValidationException ee)
            {
                return false;
            }
            return true;
        }

        void ValidationEventHandler(object sender, System.Xml.Schema.ValidationEventArgs args)
        {
            if (args.Severity == System.Xml.Schema.XmlSeverityType.Warning)
                Console.WriteLine("\tWarning: Matching schema not found.  No validation occurred." + args.Message);
            else
            {
                Console.WriteLine("\tValidation error: " + args.Message);
            }
        }

        public string BuscarCampoJson(dynamic json, string EmailProd)
        {
			log(EmailProd);
            if (json != null && EmailProd != null)
            {
                EmailBitacoraPrdDiccionario buscar = Newtonsoft.Json.JsonConvert.DeserializeObject<EmailBitacoraPrdDiccionario>(EmailProd);
                List<string> fieldNames = new List<string>();
                json = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                List<EmailBitacoraPrd> listprod = new List<EmailBitacoraPrd>();

                foreach (var input in json)
                {
                    if (input.GetType() == typeof(Newtonsoft.Json.Linq.JObject))
                    {
                        // Create JObject from object
                        Newtonsoft.Json.Linq.JObject inputJson = Newtonsoft.Json.Linq.JObject.FromObject(input);

                        // Read Properties
                        var properties = inputJson.Properties();

                        EmailBitacoraPrd prod = new EmailBitacoraPrd();

                        // Loop through all the properties of that JObject
                        foreach (var property in properties)
                        {
                            // Check if there are any sub-fields (nested)
                            // i.e. the value of any field is another JObject or another JArray
                            if (property.Value.GetType() == typeof(Newtonsoft.Json.Linq.JObject) ||
                            property.Value.GetType() == typeof(Newtonsoft.Json.Linq.JArray))
                            {

                            }
                            else
                            {
                                string field = property.Name.Replace("\n", "");
                                if (buscar.EmailBitacoraPrdItem != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdItem)
                                    {
                                        if (field == valor)
                                        {
                                            prod.EmailBitacoraPrdItem = int.Parse(property.Value.ToString());
                                        }
                                    }
                                if (buscar.EmailBitacoraPrdNombre != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdNombre)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdNombre = property.Value.ToString();
                                    }
                                if (buscar.EmailBitacoraPrdCantidad != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdCantidad)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdCantidad = decimal.Parse(property.Value.ToString().Replace(".", ","));
                                    }
                                if (buscar.EmailBitacoraPrdUndMed != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdUndMed)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdUndMed = property.Value.ToString();
                                    }
                                if (buscar.EmailBitacoraPrdMarca != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdMarca)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdMarca = property.Value.ToString();
                                    }
                                if (buscar.EmailBitacoraPrdCodigoCli != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdCodigoCli)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdCodigoCli = property.Value.ToString();
                                    }
								if (buscar.EmailBitacoraPrdNombreLargo != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdNombreLargo)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdNombreLargo = property.Value.ToString();
                                    }
								if (buscar.EmailBitacoraRFQNumber != null)
                                    foreach (var valor in buscar.EmailBitacoraRFQNumber)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraRFQNumber = property.Value.ToString();
                                    }
								if (buscar.EmailBitacoraPrdNumParte != null)
                                    foreach (var valor in buscar.EmailBitacoraPrdNumParte)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraPrdNumParte = property.Value.ToString();
                                    }
								if (buscar.EmailBitacoraReqFecha != null)
                                    foreach (var valor in buscar.EmailBitacoraReqFecha)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraReqFecha = property.Value.ToString();
                                    }
								if (buscar.EmailBitacoraReqHora != null)
                                    foreach (var valor in buscar.EmailBitacoraReqHora)
                                    {
                                        if (field == valor)
                                            prod.EmailBitacoraReqHora = property.Value.ToString();
                                    }
                            }
                        }
                        listprod.Add(prod);
                    }
                    else if (input.GetType() == typeof(Newtonsoft.Json.Linq.JValue))
                    {
                        Newtonsoft.Json.Linq.JValue inputJson = Newtonsoft.Json.Linq.JValue.FromObject(input);

                        // for direct values, there is no field name
                        fieldNames.Add(inputJson.Value.ToString());
                    }
                }
                return Newtonsoft.Json.JsonConvert.SerializeObject(listprod);
            }
            return "";
        }
		public DateTime GetFirstDateFromString(string input)
	    {
			string pattern = @"\d{2}\-\d{2}\-\d{4}";
            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(input, pattern);
            DateTime result = new DateTime();
			foreach(var value in match.Groups)  
			    if (DateTime.TryParseExact(value.ToString(), "dd-MM-yyyy", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out result))
				    return result;
			return result;
		}

        public DateTime GetFirstHourFromString(string input)
        {
            string pattern = @"\b[0-9]?\d\b\:\d{2}";
            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(input, pattern);
            DateTime result = new DateTime();
            foreach (var value in match.Groups)
                if (DateTime.TryParseExact(value.ToString(), "H:mm", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out result))
                    return result;
            return result;
        }
		public string GetPDFFromHtml(string html)
        {
            var pdf = Pdf
                        .From(html)
                        .OfSize(PaperSize.A4)
                        .WithTitle("Title")
                        .WithoutOutline()
                        .WithMargins(1.25.Centimeters())
                        .Portrait()
                        .Comressed()
                        .Content();
            return Convert.ToBase64String(pdf, 0, pdf.Length);
        }
		
        public void log(string txt)
        {
            String fecha = System.DateTime.Today.ToString("dd-MM-yyyy hh:mm:ss:ffffff");
            string logFilePath = @"C:\Logs\Log-" + System.DateTime.Today.ToString("dd-MM-yyyy") + "." + "txt";
            System.IO.File.AppendAllText(logFilePath, String.Format("{0} - {1}\r\n", fecha, txt));
        }

        public string XmlToXsd(string xml)
        {
			xml = DeleteEmptyTag(xml);
            string xsd = string.Empty;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreComments = true;        
            XmlReader reader = XmlReader.Create(new StringReader(xml), settings);
            XmlSchemaSet schemaSet = new XmlSchemaSet();
            XmlSchemaInference schema = new XmlSchemaInference();
            schemaSet = schema.InferSchema(reader);
            XmlWriterSettings settingsXSD = new XmlWriterSettings();
            settingsXSD.OmitXmlDeclaration = true;
			settingsXSD.Indent = true;
            foreach (XmlSchema s in schemaSet.Schemas())
            {
                using (var stringWriter = new StringWriter())
                {
                    using (var writer = XmlWriter.Create(stringWriter, settingsXSD))
                    {
                        s.Write(writer);
                    }

                    xsd = stringWriter.ToString();
                }
            }
           //  log(xsd);
            return xsd;
        }
		
		public string FormatHtml(string xml)
        {
			String Result = "";
            using (MemoryStream MS = new MemoryStream())            
            {
                using (XmlTextWriter W = new XmlTextWriter(MS, Encoding.Unicode))                
                {
                    XmlDocument D = new XmlDocument();
                    try
                    {
                        // Load the XmlDocument with the XML.
                        D.LoadXml(xml);
                        W.Formatting = Formatting.Indented;
                        // Write the XML into a formatting XmlTextWriter
                        D.WriteContentTo(W);
                        W.Flush();
                        MS.Flush();
                        // Have to rewind the MemoryStream in order to read
                        // its contents.
                        MS.Position = 0;
                        // Read MemoryStream contents into a StreamReader.
                        StreamReader SR = new StreamReader(MS);
                        // Extract the text from the StreamReader.
                        String FormattedXML = SR.ReadToEnd();
                        Result = FormattedXML;
                    }
                    catch (XmlException ex)
                    {
                        Result = ex.ToString();
                    }
                    W.Close();
                }
                MS.Close();
            }
            return Result;
        }

        public string PrettyXml(string xml)
        {
            var stringBuilder = new StringBuilder();

            var element = System.Xml.Linq.XElement.Parse(xml);

            var settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;
            settings.Indent = true;
            //settings.NewLineOnAttributes = true;

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                element.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }
		public string SaveEML(GxSimpleCollection<string> to, string from, string subject, string body, string emailDir, GxSimpleCollection<string> Attachment)
        {
            string FullName = string.Empty;
			int i = 0;
            System.IO.Directory.CreateDirectory(emailDir);
			string tomsg = string.Empty;
            using (var client = new System.Net.Mail.SmtpClient())
            {
				if (to.Count > 0)
				{
					tomsg = ((string)to.Item(1));
				}
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage(from, tomsg, subject, body);
				i = 2;
				while ( i <= to.Count )
				{
					string item_to = ((string)to.Item(i));
					msg.To.Add(item_to);
					i = (int)(i+1);
				}
				i = 1;
				while ( i <= Attachment.Count )
				{
					string item_Attachment = ((string)Attachment.Item(i));
					msg.Attachments.Add(new System.Net.Mail.Attachment(item_Attachment));
					i = (int)(i+1);
				}
				msg.IsBodyHtml = true;
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory;
                client.PickupDirectoryLocation = emailDir;
                try
                {
                    client.Send(msg);
                }
                catch (Exception ex)
                {
                }
                var defaultMsgPath = new DirectoryInfo(emailDir).GetFiles().OrderByDescending(f => f.LastWriteTime)
                                                               .First();
                FullName = defaultMsgPath.Name;                
            }
            return FullName;
        }
		
		 public string ConvertEMLtoMSG(string eml)
        {
            try
            {
                string msg = Path.ChangeExtension(eml, ".msg");
                MsgKit.Converter.ConvertEmlToMsg(eml, msg);
                return msg;
            }
            catch (Exception ee)
            {
                return ee.Message;
            }
        }
		public string DeleteEmptyTag(string xml)
        {
            string xslt = @"<xsl:stylesheet version=""1.0""
                                xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">
                                <xsl:output indent=""yes""/>
                                <xsl:strip-space elements=""*""/>
                                <xsl:template match=""node()|@*"">
                                    <xsl:copy>
                                        <xsl:apply-templates select=""node()|@*""/>
                                    </xsl:copy>
                                </xsl:template>
                                <xsl:template match=""*[not(@*|*|comment()|processing-instruction()) and normalize-space()='']""/>
                            </xsl:stylesheet>";

            return TransformDocument(xml, xslt);
        }
		
		public string DeleteEmptyTrTag(string xml)
        {
            string xslt = @"<xsl:stylesheet version=""1.0""
                                xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">
                                <xsl:output indent=""yes""/>
                                <xsl:strip-space elements=""*""/>
                                <xsl:template match=""node()|@*"">
                                    <xsl:copy>
                                        <xsl:apply-templates select=""node()|@*""/>
                                    </xsl:copy>
                                </xsl:template>
                                <xsl:template match=""tr[not(@*|*|comment()|processing-instruction()) and normalize-space()='']""/>
                            </xsl:stylesheet>";

            return TransformDocument(xml, xslt);
        }
		public GxSimpleCollection<string> ExtractEmails(string data)
        {
			GxSimpleCollection<string> c_correos = new GxSimpleCollection<string>() ;
            System.Text.RegularExpressions.Regex emailRegex = new System.Text.RegularExpressions.Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            System.Text.RegularExpressions.MatchCollection emailMatches = emailRegex.Matches(data);
            StringBuilder sb = new StringBuilder();
            
            foreach (System.Text.RegularExpressions.Match emailMatch in emailMatches)
            {
				if ( ! (c_correos.IndexOf(StringUtil.RTrim( emailMatch.Value))>0) )
				{
					c_correos.Add(emailMatch.Value, 0);
				}
				
            }
            return c_correos;
        }
		public string DeleteTagHtml(string html)
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(html);
            return htmlDoc.DocumentNode.InnerText;
        }
		public bool ExisteConexionInternet()
        {
            try
            {
                bool RedActiva = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
                if (RedActiva)
                {
                    System.Uri Url = new System.Uri("https://www.google.com/");
                    System.Net.WebRequest WebRequest;
                    WebRequest = System.Net.WebRequest.Create(Url);
                    System.Net.WebResponse objetoResp;
                    objetoResp = WebRequest.GetResponse();
                    objetoResp.Close();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
    public class TableHeader
    {
        public string Header { get; set; }
    }
    public class EmailBitacoraPrd
    {
        public int EmailBitacoraPrdItem { get; set; }
        public string EmailBitacoraPrdNombre { get; set; }
        public decimal EmailBitacoraPrdCantidad { get; set; }
        public string EmailBitacoraPrdUndMed { get; set; }
        public string EmailBitacoraPrdMarca { get; set; }
        public string EmailBitacoraPrdCodigoCli { get; set; }
        public string EmailBitacoraPrdNombreLargo { get; set; }
		public string EmailBitacoraRFQNumber { get; set; }
		public string EmailBitacoraPrdNumParte { get; set; }
		public string EmailBitacoraReqFecha { get; set; }
		public string EmailBitacoraReqHora { get; set; }
    }

    public class EmailBitacoraPrdDiccionario
    {
        public List<string> EmailBitacoraPrdItem { get; set; }
        public List<string> EmailBitacoraPrdNombre { get; set; }
        public List<string> EmailBitacoraPrdCantidad { get; set; }
        public List<string> EmailBitacoraPrdUndMed { get; set; }
        public List<string> EmailBitacoraPrdMarca { get; set; }
        public List<string> EmailBitacoraPrdCodigoCli { get; set; }
        public List<string> EmailBitacoraPrdNombreLargo { get; set; }
		public List<string> EmailBitacoraRFQNumber { get; set; }
		public List<string> EmailBitacoraPrdNumParte { get; set; }
		public List<string> EmailBitacoraReqFecha { get; set; }
		public List<string> EmailBitacoraReqHora { get; set; }
    }
}