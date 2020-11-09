using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace pru2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                var result = openFileDialog1.ShowDialog();
                string file = openFileDialog1.FileName;
                var workbook = new XLWorkbook(file);
                var ws1 = workbook.Worksheet(1);


                if (result == DialogResult.OK) // Test result.
                {
                    List<ReturnData> lst = new List<ReturnData>();
                    foreach (var item in ws1.Rows())
                    {


                        ReturnData data = new ReturnData();
                        data.CondicionXPATH = new List<string>();

                        var campo = item.Cell(2).Value;
                        if (campo.ToString() == "Campo") continue;
                        var codigo = item.Cell(3).Value;
                        var tipovalidacion = item.Cell(5).Value;
                        var descripcion = item.Cell(6).Value;
                        data.Nombre = descripcion.ToString();
                        data.Descripcion = descripcion.ToString();
                        data.Codigo = codigo.ToString();
                        data.Mandatorio = item.Cell(5).Value.ToString() == "Rechazo" ? true : false;
                        var pasos = item.Cell(7).Value;
                        data.MensajeError = item.Cell(8).Value.ToString();
                        var xpath = item.Cell(13).Value.ToString();
                        xpath = xpath.Substring(1);
                        string[] stringSeparators = new string[] { "\n" };
                        string[] lines = pasos.ToString().Split(stringSeparators, StringSplitOptions.None);
                        var xpathGnral = item.Cell(12).Value.ToString();
                        data.XPATH = $"{(string.IsNullOrEmpty(textBox2.Text) ? "/" : textBox2.Text)}{xpathGnral.Substring(1)}";
                        foreach (var det in lines)
                        {
                            var obj = det.Replace("\r", "");
                            switch (obj)
                            {
                                case string g when g.Contains("numeral"):
                                    data.Tipo = "En lista";
                                    data.CondicionXPATH.Add($"{textBox2.Text}{xpath}");
                                    data.CondicionXPATH.Add($"Lista {det.Substring(det.IndexOf("numeral") + 7)}");
                                    break;
                                case string f when f.Contains("Longitud"):
                                    var Longitud = Regex.Match(obj.Substring(2), @"\d+").Value;
                                    data.CondicionXPATH.Add($"string-length({textBox2.Text}{xpath})= {Longitud}");
                                    break;
                                case string h when h.Contains("Alfanumérico") || h.Contains("Alfanumperico"):
                                    break;
                                case string i when i.Contains("Numérico"):
                                    data.CondicionXPATH.Add($"number({textBox2.Text}{xpath})= number({textBox2.Text}{xpath})");
                                    break;
                                case string J when J.Contains("literal"):
                                    string text = det.Substring((det.IndexOf('“') + 1));
                                    text = text.Replace("“", "");
                                    text = text.Replace("”", "");
                                    text = text.Replace("\"", "");
                                    data.CondicionXPATH.Add($"{textBox2.Text}{xpath} = '{text}'");
                                    break;
                                case string b when b.Contains("ocurrencia"):
                                    var ocurrencia = Regex.Match(obj.Substring(2), @"\d+").Value;
                                    if (string.IsNullOrEmpty(ocurrencia))
                                    {
                                        switch (obj)
                                        {
                                            case string z when z.Contains("una"):
                                                ocurrencia = "1";
                                                break;
                                            default:
                                                ocurrencia = "1";
                                                break;
                                        }
                                    }
                                    data.CondicionXPATH.Add($"count({textBox2.Text}{xpath})= {ocurrencia}");
                                    break;
                                case string i when i.Contains("vacío"):
                                case string k when k.Contains("vacio"):
                                    data.CondicionXPATH.Add($"string-length({textBox2.Text}{xpath})> {0}");
                                    break;
                                case string m when m.Contains("vacíoG"):
                                case string n when n.Contains("vacioG"):
                                    data.CondicionXPATH.Add($"count({textBox2.Text}{xpath}/*)> {0}");
                                    break;
                                case string l when l.Contains("exista"):
                                    data.CondicionXPATH.Add($"exists({textBox2.Text}{xpath})");
                                    break;
                                default:
                                    data.CondicionXPATH.Add($"No c:{obj}");
                                    break;
                            }
                        }
                        if (string.IsNullOrEmpty(data.Tipo)) data.Tipo = "Boolean";
                        lst.Add(data);
                    }
                    textBox1.Text = JsonConvert.SerializeObject(lst);

                    string fl = "PartitionKey, RowKey, Timestamp, Active, Active@type,BreakOut,BreakOut @type, Category, Category@type,Created,Created @type, CreatedBy, CreatedBy@type,Deleted,Deleted @type, Description, Description@type,DocumentTypeCode,DocumentTypeCode @type, ErrorCode, ErrorCode@type,ErrorMessage,ErrorMessage @type, Mandatory, Mandatory@type,Name,Name @type, Priority, Priority@type,RuleId,RuleId @type, Type, Type@type,TypeListId,TypeListId @type, Updated, Updated@type,UpdatedBy,UpdatedBy @type, XPath, XPath@type,XpathEvaluation,XpathEvaluation @type";
                    var csv = new StringBuilder();
                    csv.AppendLine(fl);

                    foreach (var item in lst)
                    {
                        if (item.Tipo != "En lista")
                        {

                            var d = DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");
                            Guid g = Guid.NewGuid();
                            var data = $"\"{item.CondicionXPATH.Aggregate((s, s1) => s + " AND " + s1).ToString()} \"";
                            string line = $"new-dian-ubl21,{g},{d},true,Edm.Boolean,false,Edm.Boolean,new-dian-ubl21,Edm.String," +
                                          $"{d},Edm.DateTime,ldcolmenares@indracompany.com,Edm.String,false,Edm.Boolean,{item.Descripcion?.Replace(",", " ")}," +
                                          $"Edm.String,96,Edm.String,{item.Codigo},Edm.String,{item.MensajeError.Replace(",", " ")},Edm.String," +
                                          $"{item.Mandatorio},Edm.Boolean,{item.Descripcion?.Replace(",", " ")},Edm.String,0,Edm.Int32,{g},Edm.Guid," +
                                          $"0,Edm.Int32,,,{d},Edm.DateTime,ldcolmenares@indracompany.com,Edm.String" +
                                          $",{data}" +
                                          $",Edm.String,{item.XPATH},Edm.String";
                            csv.AppendLine(line);
                        }
                    }
                    File.WriteAllText($"{file}Result.csv", csv.ToString(), Encoding.Default);


                }
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show($"{ ex.Message}  {line.ToString()}");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var data = JsonConvert.DeserializeObject<List<ReturnData>>(textBox1.Text);
                var openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                var result = openFileDialog1.ShowDialog();
                string file = openFileDialog1.FileName;



                var docNav = XDocument.Load(file);
                XNamespace dc = "http://www.w3.org/2000/09/xmldsig#";
                XNamespace dc1 = "http://purl.org/dc/elements/1.1/";
                XNamespace dc2s = "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2";
                var nav = docNav.CreateNavigator();
                nav.SelectSingleNode("ApplicationResponse");
                foreach (var item in data)
                {
                    if (item.Tipo != "En lista")
                    {

                        item.valXPATH = (bool)nav.Evaluate(item.XPATH);
                        var det = $"\"{item.CondicionXPATH.Aggregate((s, s1) => s + " AND " + s1).ToString()} \"";
                        item.valCondicionXPATH = (bool)nav.Evaluate(det);
                    }
                    else
                    {

                        item.valXPATH = (bool)nav.Evaluate($"exists({item.XPATH})");
                    }
                }
                textBox1.Text = JsonConvert.SerializeObject(data);
            }
            catch (Exception ex)
            {

                var st = new StackTrace(ex, true);
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show($"{ ex.Message}  {line.ToString()}");
            }
        }
    }

    public class ReturnData
    {
        public string Nombre { get; set; }
        public string Descripcion { get; set; }
        public bool Mandatorio { get; set; }
        public bool Interrumpe { get; set; }
        public string Codigo { get; set; }
        public string MensajeError { get; set; }
        public string XPATH { get; set; }
        public List<string> CondicionXPATH { get; set; }
        public string Tipo { get; set; }
        public bool valXPATH { get; set; }
        public bool valCondicionXPATH { get; set; }

    }
}
