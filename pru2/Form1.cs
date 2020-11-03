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
                    List<returnData> lst = new List<returnData>();
                    foreach (var item in ws1.Rows())
                    {


                        returnData data = new returnData();
                        data.CondicionXPATH = new List<string>();

                        var campo = item.Cell(2).Value;
                        if (campo.ToString() == "Campo") continue;
                        var codigo = item.Cell(3).Value;
                        var tipovalidacion = item.Cell(5).Value;
                        var descripcion = item.Cell(6).Value;
                        data.Nombre = descripcion.ToString();
                        data.Descripcion = descripcion.ToString();
                        data.Codigo = codigo.ToString();
                        data.Interrumpe = item.Cell(5).Value.ToString() == "Rechazo" ? true : false;
                        var pasos = item.Cell(7).Value;
                        data.MensajeError = item.Cell(8).Value.ToString();
                        var xpath = item.Cell(13).Value.ToString();
                        xpath = xpath.Substring(1);
                        string[] stringSeparators = new string[] { "\n" };
                        string[] lines = pasos.ToString().Split(stringSeparators, StringSplitOptions.None);
                        data.XPATH = "{textBox2.Text}ApplicationResponse/cac:DocumentResponse/cac:Response/cbc:ResponseCode/text() = '043'";
                        foreach (var det in lines)
                        {
                            var obj = det.Replace("\r", "");
                            switch (obj)
                            {
                                case string a when a.Contains("Rechazo"):
                                    data.Interrumpe = true;
                                    break;
                                case string c when c.Contains("Opcional"):
                                    data.Mandatorio = false;
                                    break;
                                case string d when d.Contains("Obligatorio"):
                                    data.Mandatorio = true;
                                    break;
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
                                    data.CondicionXPATH.Add($"{textBox2.Text}{xpath}/text()= '{text}'");
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
                                    data.CondicionXPATH.Add($"string-length({textBox2.Text}{xpath})> {0}");
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
    }

    public class returnData
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

    }
}
