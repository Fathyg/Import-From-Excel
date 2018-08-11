using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Import_From_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataSet1.tblGroupsDataTable dtg = new DataSet1.tblGroupsDataTable();
            DataSet1.tblPartsDataTable dtp = new DataSet1.tblPartsDataTable();
            DataSet1TableAdapters.tblGroupsTableAdapter adg = new DataSet1TableAdapters.tblGroupsTableAdapter();
            DataSet1TableAdapters.tblPartsTableAdapter adp = new DataSet1TableAdapters.tblPartsTableAdapter();
            DataSet1.GroupPartsDataTable dtgp = new DataSet1.GroupPartsDataTable();
            DataSet1TableAdapters.GroupPartsTableAdapter adgp = new DataSet1TableAdapters.GroupPartsTableAdapter();
            Excel.Application xlApp = new Excel.Application();
            DialogResult dr = ofd1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                Excel.Workbook xlWbook = xlApp.Workbooks.Open(ofd1.FileName);
                string[] sheets = new string[xlWbook.Worksheets.Count];


                foreach (Excel.Worksheet item in xlWbook.Worksheets)
                {
                    if (adg.GetDataBygn(item.Name).Rows.Count <= 0)
                    {

                        adg.InsertQuery(item.Name, null);
                    }
                    int gid = adg.GetDataBygn(item.Name).First().ID;
                    int i = 3;

                    while (!string.IsNullOrEmpty(item.Range["A" + i.ToString()].Text) && !string.IsNullOrEmpty(item.Range["B" + i.ToString()].Text))
                    {
                        string a = "A" + i.ToString();
                        string b = "B" + i.ToString();
                        string c = "C" + i.ToString();
                        Excel.Range rnga = item.Range["" + a + ""];
                        Excel.Range rngb = item.Range["" + b + ""];
                        Excel.Range rngc = item.Range["" + c + ""];
                        List<string> pnk = (from p in adp.GetData() where p.PartNo == rnga.Text select p.Kind).Distinct().ToList();


                        if (!pnk.Contains(rngb.Text))
                        {
                            adp.InsertQuery(rnga.Text, rngb.Text, rngc.Text, 0, 0, 0, 0, 0, null, 0);

                        }
                        int pid = (from p in adp.GetData() where p.PartNo == rnga.Text && p.Kind == rngb.Text select p).SingleOrDefault().ID;

                        List<string> pp = (from p in adgp.GetData() where p.groupID == gid select p.partID.ToString()).ToList();
                        if (!pp.Contains(pid.ToString()))
                        {
                            adgp.InsertQuery(gid, pid, 0, null);
                        }
                        if (i >= 20 && i % 20 == 0)
                        {
                            i = i + 3;
                        }
                        else
                        {
                            i = i + 1;
                        }
                    }

                    MessageBox.Show("Successfully " + item.Name + "  " + item.Index + " has been added");
                }
                xlWbook.Close();
                xlApp.Quit();
                MessageBox.Show("تم بحمد الله");

            }



        }
    }
  
}
