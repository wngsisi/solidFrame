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
using FeatureManager = SolidWorks.Interop.sldworks.FeatureManager;


namespace solidFrame
{
    public partial class Form1 : Form
    {
        private SldWorks.SldWorks swApp;
        private ModelDoc2 swDoc;

        public Form1()
        {
            InitializeComponent();
            swApp = new SldWorks.SldWorks();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
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
                        string sheetName = "Sheet1$"; // 工作表名称
                        // 获取行数和列数
                        //string countrowQuery = $"SELECT COUNT(*) FROM [{sheetName}]";
                        //OleDbCommand countCommand = new OleDbCommand(countrowQuery, connection);
                        //int rowCount = (int)countCommand.ExecuteScalar();
                        //string counlineQuery = $"SELECT * FROM [{sheetName}]";
                        //int lineCount = (int)countCommand.ExecuteScalar();
                        // 创建一个二维数组来存储数据
                        //Console.WriteLine($"共有 ({rowCount},{lineCount}) ");
                        //string[,] data = new string[rowCount, lineCount];
                        string[,] data = new string[870, 870];
                        // 打开一个新的SolidWorks零件文档
                        swDoc = (ModelDoc2)swApp.NewDocument(@"E:\gb_part.SLDPRT", 0, 0, 0);
                        swApp.ActivateDoc2("零件1", false, 0);
                        for (int row = 1; row <= 870; row++) //索引超出数组界限 修改此处,文档为870*870
                        {
                            for (int col = 1; col <= 870; col++)
                            {
                                string cellAddress = GetExcelColumnName(col) + row;
                                string query = $"SELECT * FROM [{sheetName}{cellAddress}:{cellAddress}]";
                                using (OleDbCommand command = new OleDbCommand(query, connection))
                                {
                                    using (OleDbDataReader reader = command.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            data[row - 1, col - 1] = reader[0].ToString();
                                            double cellValue = double.Parse(data[row - 1, col - 1]);
                                            // 在SolidWorks中绘制正方形
                                            DrawSquare(row, col);
                                            //Console.WriteLine($"单元格 ({row},{col}) 绘制完成");
                                            // 开始拉伸
                                            cellValue = cellValue * 0.001; // 将单元格值转换为
                                            ExtrudeRectangle(cellValue);
                                            //Console.WriteLine($"单元格 ({row},{col}) 拉伸完成");
                                        }
                                    }
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
        }

        public void DrawSquare(int row, int col)
        {
            double size = 4.6*0.001; // 正方形的边长
            double x = col * size; // 根据列索引确定x坐标
            double y = row * size; // 根据行索引确定y坐标
            swDoc.SketchManager.InsertSketch(true);
            swDoc.SketchManager.CreateCornerRectangle(x, y, 0, x + size, y + size, 0);
            swDoc.ClearSelection2(true);
        }

        public void ExtrudeRectangle(double length)

        {
            FeatureManager featureManager = swDoc.FeatureManager;
            bool reverseDirection = false;
            if (length < 0)
            {
                reverseDirection = true;
                length = -length;
                return;
            }

            
            featureManager.FeatureExtrusion2(true, true , reverseDirection, 0, 0, length, 0, false, false, true , false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
        }

        private string GetExcelColumnName(int columnNumber)// 将Excel列号转换为Excel列名
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private void button2_Click(object sender, EventArgs e)//测试按钮
        {
            // 打开一个新的SolidWorks零件文档
            swDoc = (ModelDoc2)swApp.NewDocument(@"E:\gb_part.SLDPRT", 0, 0, 0);
            swApp.ActivateDoc2("零件1", false, 0);

            // 绘制一个正方形
            DrawSquare(1, 1);
            Console.WriteLine("单元格 (1,1) 绘制完成");

            // 拉伸矩形为实体
            double cell = -0.05; // 手动指定拉伸长度
            ExtrudeRectangle(cell);
            Console.WriteLine($"矩形已拉伸为实体，长度为: {cell}");
        }
    }
}
