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
            string filePath = "path_to_your_csv_file.csv"; // 请替换为你的CSV文件路径
            OpenFileDialog ofd = new OpenFileDialog();
            FileDialog fileDialog = null;
            //ofd.Restorehistory * false;
            ofd.Multiselect = true;
            ofd.ShowDialog();
            st = ofd.FileNames;
            stl = ofd.SafeFileNames;

            ss = System.IO.Path.GetFileName(ofd.FileName);//获取文件名
            ss2 =System.IO.Path.GetFullPath(ofd.FileName);//获取文件路径
            SolidWorks.Interop.swconst.swDocumentTypes_e swDocType = SolidWorks.Interop.swconst.swDocumentTypes_e.
            int errors = 0;
            int warnings = 0;
            ModelDoc2 swModel;


        }

        private void CreateCube(double edgeLength)
        {
            // 这里可以实现立方体的创建逻辑
            // C# 示例代码：创建一个简单的立方体
            SolidWorks.Interop.swconst.swDocumentTypes_e swDocType = SolidWorks.Interop.swconst.swDocumentTypes_e.swDoc3D钣金;
            int errors = 0;
            int warnings = 0;
            ModelDoc2 swModel;
            try
            {
                swModel = swApp.NewDocument("C:\\ProgramData\\SolidWorks\\SOLIDWORKS 2022\\templates\\Default.SLDPRT", 0, 0, 0);
                PartDoc swPart = (PartDoc)swModel;
                Feature swFeature = swPart.FeatureManager.FeatureExtrusion2(false, false, false, null, null, false, false, 0, 0, 100, 0, 0, true, true, false, 0, 0, false, false, false, 0, ref errors, ref warnings);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            // 例如，显示立方体的边长和体积
            double volume = Math.Pow(edgeLength, 3);
            MessageBox.Show($"立方体的边长: {edgeLength}\n立方体的体积: {volume}");
        }
    }
}