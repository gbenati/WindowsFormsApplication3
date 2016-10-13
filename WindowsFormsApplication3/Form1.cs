using System;
using System.IO;
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
        internal List<Group> GroupList;
        internal List<CalibrationScaling> CScalingList;

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
            int RetVal = 0 ;
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

                    // If it is a CalibrationValue
                    if (RetVal == 0)
                    {
                        SignalValue1.upload(ref XLSECTSignal1, false, 0);
                        SignalValue1.AppendToFile(ref fileXML);
                        //                        SignalValue1.Show();
                    }
                    else
                    {
                        if (RetVal > 1)
                        {
                            for (index = 0; index < RetVal; index++)
                            {
                                SignalValue1.upload(ref XLSECTSignal1, true, index);
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

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog2_FileOk_1(object sender, CancelEventArgs e)
        {
            filePath = openFileDialog2.FileName;
//            MessageBox.Show(filePath);
            textBox2.Text = filePath;
            label1.Text = filePath;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            XMLfilePath = openFileDialog1.FileName;
//            MessageBox.Show(filePath);
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
    }
}
