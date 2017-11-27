using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace interviewmapping
{
    public partial class Form1 : Form
    {
        List<Company> allComp = new List<Company>();
        List<Student> allStud = new List<Student>();
        System.Data.DataTable dtb = new System.Data.DataTable();
        
        Dictionary<string, string> smapTable = new Dictionary<string, string>();
        Dictionary<string, string> cmapTable = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            readData();
            dtb.Columns.Add(new DataColumn("student",Type.GetType("System.String")));
            dtb.Columns.Add(new DataColumn("company", Type.GetType("System.String")));
            dtb.Columns.Add(new DataColumn("class", Type.GetType("System.String")));
            doMapping();
        }


        private void readData()
        {
            //read company data
            string FilePath = "e://v.csv";
            StreamReader sr = new StreamReader(FilePath);
            string line = "";
            while ((line = sr.ReadLine()) != null)
            {
                string[] row = line.Split(',');
                Company com = new Company();
                com.Name = row[0].Trim();
                    for (int i = 1; i < row.Length; i++)
                    {
                        com.CandidateStudent.Add(row[i].Trim());
                    }
                        
                allComp.Add(com);
            }
            sr.Close();

            //read student data
            string ExcelBookFileName = "e:篩選結果.xlsx";

            Microsoft.Office.Interop.Excel.Application ExcelApp
              = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Visible = false;
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing);
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);
            for (int s = 2; s <= 40; s++)
            {
                Range range = ExcelApp.get_Range("A"+s.ToString(), Type.Missing);
                Student stud = new Student();
                stud.Name = range.Text.Trim();
                stud.Class = (range.Cells[1, 2] as Range).Text;
                for (int c = 3; c <= 22; c++)
                {
                    Range rng = (range.Cells[1, c]);
                    if (!string.IsNullOrEmpty(rng.Text))
                    {
                        //stud.MappingCompany.Add(int.Parse(rng.Text), (ws1.Cells[1, c] as Range).Text );
                        stud.MappingCompany.Add(c, (ws1.Cells[1, c] as Range).Text);
                    }
                }
                allStud.Add(stud);
            }
            wb.Close(false, Type.Missing, Type.Missing);
            ExcelApp.Quit();
        }

        private void doMapping()
        {
            foreach(Student s in allStud)
            {
                foreach (var c in s.MappingCompany)
                {
                    var com = allComp.Where((e) => { return (c.Value.StartsWith(e.Name) && e.CandidateStudent.Contains(s.Name.Substring(0,3))) ; }).ToList();
                    foreach (Company cm in com)
                    {
                        DataRow dr = dtb.NewRow();
                        dr[0] = s.Name;
                        //dr[1] = cm.Name + "(" +c.Key + ")";
                        dr[1] = cm.Name;
                        dr[2] = s.Class;
                        dtb.Rows.Add(dr);
                    }
                        
                }
            }
            DataView dv = new DataView(dtb);
            dataGridView1.DataSource = dv;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            doMapping();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string ExcelBookFileName = "e:面試時程.xlsx";
            Microsoft.Office.Interop.Excel.Application ExcelApp
              = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Visible = false;
            //Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing);
            Workbook wb = ExcelApp.Workbooks.Open(Filename:ExcelBookFileName);
            Worksheet studentws = wb.Sheets["各同學面試時程"];
            Worksheet companyws = wb.Sheets["各公司面試時程"];
            studentws.Select(Type.Missing);
            for (int s = 2; s <= 40; s++)
            {
                Range stud = (studentws.Cells[1, s] as Range);
                if (stud.EntireColumn.Hidden) continue;
                Range time_range = studentws.get_Range("A1", "A104");
                
                string posx = GetExcelColumnName(stud.Column) + "1";
                string posy = GetExcelColumnName(stud.Column) + "104";
                Range sche_range = studentws.get_Range(posx, posy);
                Range classroom = companyws.get_Range("A1");
                Worksheet newws = wb.Sheets.Add();
                newws.Name = stud.Text;
                time_range.Copy(newws.Range["A1","A104"]);
                sche_range.Copy(newws.Range["B1", "B104"]);
                
                classroom.Copy(newws.Range["D2"]);
                newws.Range["D2", "N20"].Merge(false);
                //wb.Sheets["各同學面試時程"].Delete();
                //wb.Sheets["各公司面試時程"].Delete();
                //wb.SaveAs(Filename: "e:\\studemail\\" + stud.Text);
                var workbook = ExcelApp.Workbooks.Add(1);
                newws.Copy(workbook.Sheets[1]);
                workbook.SaveAs(Filename: "e:\\studemail\\" + stud.Text);
                workbook.Close();
            }
            wb.Close(false, Type.Missing, Type.Missing);
            ExcelApp.Quit();
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
