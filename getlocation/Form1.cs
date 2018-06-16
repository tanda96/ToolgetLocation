using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using ClosedXML.Excel;

namespace getlocation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string removeStringNotUse(string text)
        {

            string pattern = @"\/\*\*\/.*\( |\)";
            string input = text;
            RegexOptions options = RegexOptions.Multiline;

            string result =  Regex.Replace(input, pattern,"", RegexOptions.IgnoreCase);

            return result;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            List<String> array_location = new List<string>();

            string fileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\" + "Source.xlsx";

            var workbook = new XLWorkbook(fileName);
            var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();

            foreach (var dataRow in nonEmptyDataRows)
            {
                //for row number check
                if (dataRow.RowNumber() <500)
                {
                    //to get column # 3's data
                    var cell = dataRow.Cell(1).Value;
                    array_location.Add(cell.ToString());
                }
            }
           

            /* Viet file */
            
            string endfileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\" + "EndSource.xlsx";
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[2] {
                new DataColumn("Dia Chi", typeof(string)),
                new DataColumn("Location",typeof(string))
            });

            
            foreach (var item in array_location)
            {
                string location = item;
                String url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + location + "&key=AIzaSyBTuRonImBIawxKxL1o251SzqlZk-HG4Vg";

                WebClient wc = new WebClient();
                var json = "";
                try
                {
                    json = wc.DownloadString(url);
                }
                catch (Exception)
                {
                    MessageBox.Show("Nhap ID Group | KIểm Tra Internet | Sai Token");
                    return;
                }
                string temp = removeStringNotUse(json);
                JObject data = new JObject();
                try
                {
                    data = JObject.Parse(temp);
                }
                catch (Exception)
                {
                    Console.WriteLine("Bi Loi Json");
                }
                //MessageBox.Show(data["status"].ToString());

                if (data["status"].ToString() == "OK")
                {
                    string lat = data["results"][0]["geometry"]["location"]["lat"].ToString();
                    string lng = data["results"][0]["geometry"]["location"]["lng"].ToString();
                    Console.WriteLine(item.ToString() + " : " + lat + "," + lng);
                    dt.Rows.Add(item.ToString(), lat + "," + lng);

                }
                else
                {
                   // MessageBox.Show("Meo ra"+ item.ToString() );
                    dt.Rows.Add(item.ToString(), "Error.");
                }
            }
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Customers");

            wb.SaveAs(endfileName);
            MessageBox.Show("Xong");
           

        }
    }
}
