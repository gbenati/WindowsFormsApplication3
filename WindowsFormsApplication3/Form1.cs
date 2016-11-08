using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//using Characteristic;


namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        string filePath = @"d:\Software\XML_ECT\LIE00_ModelName.xls";
        string XMLfilePath = @"d:\Software\XML_ECT\LIE00PARTIAL.xml";
        string BasefilePath = @"d:\Software\XML_ECT\LIE00V12BASE.xml";
        string filePathDir = @"d:\Software\XML_ECT";
        private string A2LfilePath = @"d:\Software\XML_ECT\A2L\LIE00_Partial.a2l";
        private string A2LINIfilePath = @"d:\Software\XML_ECT\A2L\LIE00.ini";
        string OBJfilePath = "";
        internal List<Group> GroupList;
        internal List<CalibrationScaling> CScalingList;
        internal Cont Containr;

        public Form1()
        {
            InitializeComponent();
            textBox2.Text = filePath;
            textBox1.Text = XMLfilePath;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = filePathDir;
            openFileDialog2.ShowDialog();
        }

        private void CopyBaseFileToTarget()
        {
            string line;
            System.IO.StreamReader fileBase;
            System.IO.StreamWriter fileXML;

            fileBase = new System.IO.StreamReader(BasefilePath);
            fileXML = new System.IO.StreamWriter(XMLfilePath);

            line = fileBase.ReadLine();

            if (line != null)
            {
                do
                {
                    fileXML.WriteLine(line);
                    line = fileBase.ReadLine();
                } while (line != null);
            }
            fileBase.Close();
            fileXML.Close();
        }

        private void CopyFileAToB(string A, string B)
        {
            string line;
            System.IO.StreamReader fileA;
            System.IO.StreamWriter fileB;

            fileA = new System.IO.StreamReader(A);
            fileB = new System.IO.StreamWriter(B, true);

            line = fileA.ReadLine();

            if (line != null)
            {
                do
                {
                    line = "\t\t" + line;
                    fileB.WriteLine(line);
                    line = fileA.ReadLine();
                } while (line != null);
            }
            fileA.Close();
            fileB.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet_Parameters;
            Excel.Worksheet xlWorkSheet_Signals;
            Excel.Worksheet xlWorkSheet_Defines;
            Excel.Worksheet xlWorkSheet_States;
            string Line;
            int LineNum;
            int RetVal = 0;
            int index;
            bool ScalingExists = false;
            bool GroupExists = false;

            XLSECTParameter XLSECTParameter1 = new XLSECTParameter();
            XLSECTSignal XLSECTSignal1 = new XLSECTSignal();
            SignalValue SignalValue1 = new SignalValue();
            CalibrationValue CalibrationValue1 = new CalibrationValue();
            CalibrationSharedAxis CalibrationSharedAxis1 = new CalibrationSharedAxis();
            CalibrationCurve CalibrationCurve1 = new CalibrationCurve();
            CalibrationMap CalibrationMap1 = new CalibrationMap();
            System.IO.StreamWriter fileXML;
            /* Scaling IDs */
            CalibrationScaling CScalingNew = new CalibrationScaling();
            //            CalibrationScaling CS;
            Group GroupNew = new Group();


            GroupList = new List<Group>();
            CScalingList = new List<CalibrationScaling>();

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet_Parameters = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet_Signals = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlWorkSheet_Defines = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            xlWorkSheet_States = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

            if (File.Exists(BasefilePath))
            {
                // Copy the XML base file as a header for the new file

                CopyBaseFileToTarget();

                fileXML = new System.IO.StreamWriter(XMLfilePath, true);

                // Generate the signals and write them in the new XML file
                Line = "2";
                LineNum = 2;
                while (-2 != RetVal)
                {
                    // Elaborate the signals sheet
                    RetVal = XLSECTSignal1.upload(ref xlWorkSheet_Signals, Line);

                    // Build the Scaling IDs list 

                    CScalingNew.upload(ref XLSECTSignal1);
                    ScalingExists = false;

                    foreach (CalibrationScaling CS in CScalingList)
                    {
                        if (CS.ID == CScalingNew.ID)
                        {
                            ScalingExists = true;
                        }
                    }

                    if (ScalingExists == false)
                    {
                        CScalingList.Add(CScalingNew);
                        CScalingNew = new CalibrationScaling();
                    }

                    // Build the Group IDs list 
                    GroupNew.upload(ref XLSECTSignal1);
                    GroupExists = false;

                    foreach (Group GP in GroupList)
                    {
                        if (GP.ID == GroupNew.ID)
                        {
                            GroupExists = true;
                        }
                    }

                    if (GroupExists == false)
                    {
                        GroupList.Add(GroupNew);
                        GroupNew = new Group();

                    }

                    // If it is a Signal (Channel in ECT XML nomenclature)
                    if (RetVal == 0)
                    {
                        SignalValue1.upload(ref XLSECTSignal1, false, 0, ref Containr);
                        SignalValue1.AppendToFile(ref fileXML);
                        //                        SignalValue1.Show();
                    }
                    else
                    {
                        if (RetVal > 1)
                        {
                            for (index = 0; index < RetVal; index++)
                            {
                                SignalValue1.upload(ref XLSECTSignal1, true, index, ref Containr);
                                SignalValue1.AppendToFile(ref fileXML);
                            }
                        }
                    }
                    LineNum++;
                    Line = Convert.ToString(LineNum);
                }

                // Generate the calibrations and write them in the new XML file
                Line = "2";
                LineNum = 2;
                RetVal = 0;
                while (-2 != RetVal)
                {
                    // Elaborate the parameters sheet
                    RetVal = XLSECTParameter1.upload(ref xlWorkSheet_Parameters, Line);

                    CScalingNew.upload(ref XLSECTParameter1);
                    ScalingExists = false;

                    foreach (CalibrationScaling CS in CScalingList)
                    {
                        if (CS.ID == CScalingNew.ID)
                        {
                            ScalingExists = true;
                        }
                    }

                    if (ScalingExists == false)
                    {
                        CScalingList.Add(CScalingNew);
                        CScalingNew = new CalibrationScaling();
                    }

                    // Build the Group IDs list 
                    GroupNew.upload(ref XLSECTParameter1);
                    GroupExists = false;

                    foreach (Group GP in GroupList)
                    {
                        if (GP.ID == GroupNew.ID)
                        {
                            GroupExists = true;
                        }
                    }

                    if (GroupExists == false)
                    {
                        GroupList.Add(GroupNew);
                        GroupNew = new Group();

                    }

                    // If it is a CalibrationValue
                    if (RetVal == 0)
                    {
                        CalibrationValue1.upload(ref XLSECTParameter1);
                        CalibrationValue1.AppendToFile(ref fileXML);
                        //                        CalibrationValue1.Show();
                    }

                    // If it is a CalibrationSharedAxis
                    if (RetVal == 1)
                    {
                        CalibrationSharedAxis1.upload(ref XLSECTParameter1);
                        CalibrationSharedAxis1.AppendToFile(ref fileXML);
                        //                        CalibrationSharedAxis1.Show();
                    }

                    // If it is a CalibrationCurve
                    if (RetVal == 2)
                    {
                        CalibrationCurve1.upload(ref XLSECTParameter1);
                        CalibrationCurve1.AppendToFile(ref fileXML);
                        //                        CalibrationCurve1.Show();
                    }

                    // If it is a CalibrationMap
                    if (RetVal == 3)
                    {
                        CalibrationMap1.upload(ref XLSECTParameter1);
                        CalibrationMap1.AppendToFile(ref fileXML);
                        //                        CalibrationMap1.Show();
                    }

                    LineNum++;
                    Line = Convert.ToString(LineNum);

                }
                foreach (Group GP in GroupList)
                {
                    GP.AppendToFile(ref fileXML);
                }

                foreach (CalibrationScaling CS in CScalingList)
                {
                    CS.AppendToFile(ref fileXML);
                }

                fileXML.WriteLine("</LIE00V12PARTIAL>");
                fileXML.Close();

            }
            else
            {
                MessageBox.Show(" Base File doesn't exist");
            }


            /*

                        MessageBox.Show(
                            xlWorkSheet_Parameters.get_Range("A2", "A2").Value2.ToString() + "\r\n" +
                            xlWorkSheet_Parameters.get_Range("A3", "A3").Value2.ToString() + "\r\n" +
                            xlWorkSheet_Parameters.get_Range("A4", "A4").Value2.ToString() + "\r\n" +
                            xlWorkSheet_Parameters.get_Range("A5", "A5").Value2.ToString() + "\r\n" +
                            xlWorkSheet_Parameters.get_Range("A6", "A6").Value2.ToString() + "\r\n" 
                                       );
            */
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet_Parameters);
            releaseObject(xlWorkSheet_Signals);
            releaseObject(xlWorkSheet_Defines);
            releaseObject(xlWorkSheet_States);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            MessageBox.Show("Finished!");

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog2_FileOk_1(object sender, CancelEventArgs e)
        {
            filePath = openFileDialog2.FileName;
            textBox2.Text = filePath;
            label1.Text = filePath;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            XMLfilePath = openFileDialog1.FileName;
            textBox1.Text = XMLfilePath;
            label1.Text = filePath;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = XMLfilePath;
            openFileDialog1.ShowDialog();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog3.InitialDirectory = OBJfilePath;
            openFileDialog3.ShowDialog();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog3_FileOk(object sender, CancelEventArgs e)
        {
            Process p;
            ProcessStartInfo p_info;

            //            System.IO.StreamReader fileBase;
            StreamWriter filedump;
            StreamWriter filebss;
            StreamWriter filecalib;
            //            string output = null ;
            string line = null;
            StreamReader SR;
            RawSymbol C;
            Symbol S;
            string dumpline;

            Containr = new Cont();

            filedump = new System.IO.StreamWriter(@"d:/tmp/objdump.temp");
            filebss = new System.IO.StreamWriter(@"d:/tmp/objdump.bss");
            filecalib = new System.IO.StreamWriter(@"d:/tmp/objdump.calib");

            /* Retrieve filename */
            OBJfilePath = openFileDialog3.FileName;
            textBox3.Text = OBJfilePath;

            /* Create symbol tabel using objdump */
            p = new Process();

            p_info = p.StartInfo;
            p_info.RedirectStandardOutput = true;
            p_info.UseShellExecute = false;
            p_info.FileName = "objdump.exe";
            p_info.Arguments = "-x " + OBJfilePath;

            p.Start();

            SR = p.StandardOutput;
            line = SR.ReadLine();

            if (line != null)
            {
                do
                {
                    line = SR.ReadLine();

                    C = new RawSymbol(ref line);
                    S = new Symbol(ref C);

                    if (S.section == ".bss")
                    {
                        S.section = "bss";
                        Containr.SymbolBssList.Add(S);
                        dumpline = S.ConvertToLine();
                        if (dumpline != null) filebss.WriteLine(dumpline);
                    }
                    else if (S.section == ".calibu")
                    {
                        S.section = "cal";
                        Containr.SymbolCalList.Add(S);
                        dumpline = S.ConvertToLine();
                        if (dumpline != null) filecalib.WriteLine(dumpline);
                    }

                    dumpline = S.ConvertToLine();
                    if (dumpline != null) filedump.WriteLine(dumpline);

                } while (line != null);

            }
            p.WaitForExit();

            //            MessageBox.Show(output);
            filedump.Close();
            filebss.Close();
            filecalib.Close();
            MessageBox.Show("Symbol table created");

            Containr.SymbolTableExists = true;

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog5.InitialDirectory = A2LINIfilePath;
            openFileDialog5.ShowDialog();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            int i;
            string line;
            string section;
            string[] MOD_PARfilePath = new string[0];
            string[] A2MLfilePath = new string[0];
            string[] MOD_COMMONfilePath = new string[0];
            string[] IF_DATAfilePath = new string[0];
            bool section_done_mod_par = false;
            bool section_done_a2ml = false;
            bool section_done_mod_common = false;
            bool section_done_if_data = false;

            Char[] delimiter = { ';' };

            System.IO.StreamReader fileINI;
            System.IO.StreamWriter fileA2L;

            Directory.SetCurrentDirectory(Path.GetDirectoryName(A2LfilePath));

            // Generate A2L file

            fileA2L = new System.IO.StreamWriter(A2LfilePath);

            // Generate the header
            // Write the header to the A2L file
            line = "/* --------------------------------------------------------*/";
            fileA2L.WriteLine(line);
            line = "/* ASAP2 file created by Data Dictionary Wizard            */";
            fileA2L.WriteLine(line);
            line = "/* Created: "+DateTime.Now.ToString() + "                            */";
            fileA2L.WriteLine(line);
            line = "/* Benati Ltd (C)  2016                                    */";
            fileA2L.WriteLine(line);
            line = "/* http://www.benati.co.uk                                 */";
            fileA2L.WriteLine(line);
            line = "/* --------------------------------------------------------*/";
            fileA2L.WriteLine(line);
            line = " ";
            fileA2L.WriteLine(line);
            line = "ASAP2_VERSION 1 51";
            fileA2L.WriteLine(line);
            line = "/begin PROJECT LIE00V12AVENTADOR \"E00117 Lamborghini V E0011700\"";
            fileA2L.WriteLine(line);
            line = "\t/begin HEADER \"\"";
            fileA2L.WriteLine(line);
            line = "\t\tVERSION \"E00117 Lamborghini V E0011700\"";
            fileA2L.WriteLine(line);
            line = "\t/end HEADER ";
            fileA2L.WriteLine(line);
            line = "\t/begin MODULE E0011700 \"E00117 Lamborghini V E0011700\"";
            fileA2L.WriteLine(line);
            line = " ";
            fileA2L.WriteLine(line);

            // Close the file because it will be reopen later
            fileA2L.Close();

            // Open INI file, get file names

            fileINI = new System.IO.StreamReader(A2LINIfilePath);

            do
            {
                do
                {
                    line = fileINI.ReadLine();

                } while (((line != "[MOD_PAR]") && (line != "[A2ML]") && (line != "[MOD_COMMON]") && (line != "[IF_DATA]")) && (false == fileINI.EndOfStream));

                section = line;

                do
                {
                    line = fileINI.ReadLine();

                } while (  (false == line.StartsWith("SOURCE=")) 
                         &&(false == fileINI.EndOfStream)
                        );

                switch (section)
                {
                    case "[MOD_PAR]":
                        MOD_PARfilePath = line.Substring(7).Split(delimiter, System.StringSplitOptions.RemoveEmptyEntries);
                        section_done_mod_par = true;
                        break;
                    case "[A2ML]":
                        A2MLfilePath = line.Substring(7).Split(delimiter, System.StringSplitOptions.RemoveEmptyEntries);
                        section_done_a2ml = true;
                        break;
                    case "[MOD_COMMON]":
                        MOD_COMMONfilePath = line.Substring(7).Split(delimiter, System.StringSplitOptions.RemoveEmptyEntries);
                        section_done_mod_common = true;
                        break;
                    case "[IF_DATA]":
                        IF_DATAfilePath = line.Substring(7).Split(delimiter, System.StringSplitOptions.RemoveEmptyEntries);
                        section_done_if_data = true;
                        break;
                    default:
                        break;

                }
            } while (   (false == section_done_if_data) 
                     || (false == section_done_mod_common)
                     || (false == section_done_a2ml)
                     || (false == section_done_mod_par) );
#if CAZ
            if (  (MOD_PARfilePath.Length != 0)
                && (A2MLfilePath.Length != 0)
                && (MOD_COMMONfilePath.Length != 0)
                && (IF_DATAfilePath.Length != 0)
                )
            {
                MessageBox.Show("MOD_PARfilePath = " + MOD_PARfilePath[0] + "\n" + "A2MLfilePath = " + A2MLfilePath[0] + "\n" + "MOD_COMMONfilePath = " + MOD_COMMONfilePath[0] + "\n" + "IF_DATAfilePath = " + IF_DATAfilePath[0] + "\n");
            }
#endif
            // Copy A2ML to the A2L file
            for (i = 0; i < A2MLfilePath.Length; i++) CopyFileAToB(A2MLfilePath[i], A2LfilePath);
            // Copy MOD_PAR to the A2L file
            for (i = 0; i < MOD_PARfilePath.Length; i++) CopyFileAToB(MOD_PARfilePath[i], A2LfilePath);
            // Copy MOD_COMMON to the A2L file
            for (i = 0; i < MOD_COMMONfilePath.Length; i++) CopyFileAToB(MOD_COMMONfilePath[i], A2LfilePath);
            // Copy IF_DATA to the A2L file
            for (i = 0; i < IF_DATAfilePath.Length; i++) CopyFileAToB(IF_DATAfilePath[i], A2LfilePath);

            // Write footer
            fileA2L = new System.IO.StreamWriter(A2LfilePath,true); // Append

            // Generate the footer
            // Write the footer to the A2L file
            line = "    /end MODULE";
            fileA2L.WriteLine(line);
            line = "/end PROJECT";
            fileA2L.WriteLine(line);

            // Close the file because it will be reopen later
            fileA2L.Close();

            MessageBox.Show("A2L file generated ");
        }

        private void openFileDialog4_FileOk(object sender, CancelEventArgs e)
        {
            // Store A2L filel name
            A2LfilePath = openFileDialog4.FileName;
            textBox4.Text = A2LfilePath;
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            openFileDialog4.InitialDirectory = A2LfilePath;
            openFileDialog4.ShowDialog();

        }

        private void openFileDialog5_FileOk(object sender, CancelEventArgs e)
        {
            // Store and show A2L INI file name
            A2LINIfilePath = openFileDialog5.FileName;
            textBox5.Text = A2LINIfilePath;
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
