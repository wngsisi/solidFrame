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
            
            OpenFileDialog ofd = new OpenFileDialog();
            FileDialog fileDialog = null;
            //ofd.Restorehistory * false;
            ofd.Multiselect = true;
            ofd.ShowDialog();
            st = ofd.FileNames;
            stl = ofd.SafeFileNames;

            ss = System.IO.Path.GetFileName(ofd.FileName);//获取文件名
            ss2 = System.IO.Path.GetFullPath(ofd.FileName);//获取文件路径
            // 修复 CS0117 错误
            SolidWorks.Interop.swconst.swDocumentTypes_e swDocType = SolidWorks.Interop.swconst.swDocumentTypes_e.swDocPART;
            int errors = 0;
            int warnings = 0;
            ModelDoc2 swModel;

            // 使用 errors 变量以修复 CS0219 错误
            if (errors == 0)
            {
                // 处理没有错误的情况
            }
        }

    
}
}
