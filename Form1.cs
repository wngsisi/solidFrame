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

using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldworks;
using Microsoft.Win32;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using FileDialog = System.Windows.Forms.FileDialog;
using static System.Net.WebRequestMethods;
using SldWorks;
using ModelDoc2 = SolidWorks.Interop.sldworks.ModelDoc2;
using System.Data.OleDb;



namespace solidFrame
{
    public partial class Form1 : Form
    {
        public string[] st;
        public string[] stl;
        public string ss;
        public string ss2;
        public string ss3;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            FileDialog fileDialog = null;
            //ofd.Restorehistory * false;
            ofd.Multiselect = true;
            //ofd.ShowDialog();


            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=NO;'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        string sheetName = "Sheet1$"; // 修改为你的工作表名称
                        int row = 1; // 行索引
                        int col = 1; // 列索引
                        string cellAddress = GetExcelColumnName(col) + row;
                        string query = $"SELECT * FROM [{sheetName}{cellAddress}:{cellAddress}]";

                        using (OleDbCommand command = new OleDbCommand(query, connection))
                        {
                            using (OleDbDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string cellValue = reader[0].ToString();
                                    MessageBox.Show($"单元格 ({row},{col}) 的值是: {cellValue}");
                                    Console.WriteLine($"单元格 ({row},{col}) 的值是: {cellValue}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"读取Excel文件时出错: {ex.Message}");
                    }
                }
            }
            //SolidWorks.Interop.swconst.swDocumentTypes_e swDocType = SolidWorks.Interop.swconst.swDocumentTypes_e.swDocPART;
            int errors = 0;
            int warnings = 0;
            ModelDoc2 swModel;

            // 使用 errors 变量以修复 CS0219 错误
            if (errors == 0)
            {
                // 处理没有错误的情况
            }
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            Console.WriteLine("get函数被执行");
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}       
