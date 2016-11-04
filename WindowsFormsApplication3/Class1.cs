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

namespace WindowsFormsApplication3
{
    public class Group
    {
        /*
            <Groups>
            <ID>SENSORS_LINEARISATION</ID>
            <VarNameL0>SENSORS_LINEARISATION</VarNameL0>
            <VarNameL1>SENSORS_LINEARISATION</VarNameL1>
            <DescL0 />
            <DescL1 />
            <Visible>true</Visible>
            </Groups>
         */

        internal string ID;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;
        internal string Visible;

        public Group()
        {
            ID = "    <ID>GROUP_NAME</ID>";
            VarNameL0 = "    <VarNameL0>GROUP_NAME</VarNameL0>";
            VarNameL1 = "    <VarNameL1>GROUP_NAME</VarNameL1>";
            DescL0 = "    <DescL0 />";
            DescL1 = "    <DescL1 />";
            Visible = "    <Visible>true</Visible>";
        }
        public int upload(ref XLSECTSignal P)
        {
            ID = "    <ID>"+P.sigSource + "</ID>";
            VarNameL0 = "    <VarNameL0>" + P.sigSource + "</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.sigSource + "</VarNameL1>";
            DescL0 = "    <DescL0 />";
            DescL1 = "    <DescL1 />";
            Visible = "    <Visible>true</Visible>";

            return (0);
        }
        public int upload(ref XLSECTParameter P)
        {
            ID = "    <ID>" + P.parSource + "</ID>";
            VarNameL0 = "    <VarNameL0>" + P.parSource + "</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.parSource + "</VarNameL1>";
            DescL0 = "    <DescL0 />";
            DescL1 = "    <DescL1 />";
            Visible = "    <Visible>true</Visible>";

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <Groups>");
            fileXML.WriteLine(ID);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);
            fileXML.WriteLine(Visible);
            fileXML.WriteLine("  </Groups>");

            return (0);
        }
    }

    public class CalibrationScaling
    {
        /*
            <CalibrationScalings>
            <ID>S_1_1_0_0_1_4</ID>
            <Type>1</Type>
            <LimitType />
            <LimitMode>0</LimitMode>
            <RangeCorrectionModeForEquiSpacedBP>0</RangeCorrectionModeForEquiSpacedBP>
            <EquiSpacedBPCount>0</EquiSpacedBPCount>
            <G1>1</G1>
            <O1>0</O1>
            <G2>0</G2>
            <O2>1</O2>
            </CalibrationScalings>

        */
        internal string ID;
        internal string CVType;
        internal string LimitType;
        internal string LimitMode;
        internal string RangeCorrectionModeForEquiSpacedBP;
        internal string EquiSpacedBPCount;
        internal string G1;
        internal string O1;
        internal string G2;
        internal string O2;

        public CalibrationScaling()
        {
            ID = "    <ID>S_1_1_0_0_1</ID>";
            CVType = "    <Type>1</Type>";
            LimitType = "    <LimitType />";
            LimitMode = "    <LimitMode>0</LimitMode>";
            RangeCorrectionModeForEquiSpacedBP = "    <RangeCorrectionModeForEquiSpacedBP>0</RangeCorrectionModeForEquiSpacedBP>";
            EquiSpacedBPCount = "    <EquiSpacedBPCount>0</EquiSpacedBPCount>";
            G1 = "    <G1>1</G1>";
            O1 = "    <O1>0</O1>";
            G2 = "    <G2>0</G2>";
            O2 = "    <O2>1</O2>";
        }

        public int upload(ref XLSECTSignal P)
        {
            string scaling_string = "S_1_1_0_0_1";

            scaling_string = "S_1_" + Convert.ToString(P.sigKa) + "_" + Convert.ToString(P.sigKb) + "_" + Convert.ToString(P.sigKc) + "_" + Convert.ToString(P.sigKd);

            ID = "    <ID>"+ scaling_string + "</ID>";
            CVType = "    <Type>1</Type>";
            LimitType = "    <LimitType />";
            LimitMode = "    <LimitMode>0</LimitMode>";
            RangeCorrectionModeForEquiSpacedBP = "    <RangeCorrectionModeForEquiSpacedBP>0</RangeCorrectionModeForEquiSpacedBP>";
            EquiSpacedBPCount = "    <EquiSpacedBPCount>0</EquiSpacedBPCount>";
            G1 = "    <G1>" + Convert.ToString(P.sigKa) + "</G1>";
            O1 = "    <O1>" + Convert.ToString(P.sigKb) + "</O1>";
            G2 = "    <G2>" + Convert.ToString(P.sigKc) + "</G2>";
            O2 = "    <O2>" + Convert.ToString(P.sigKd) + "</O2>";

            return (0);
        }
        public int upload(ref XLSECTParameter P)
        {
            string scaling_string = "S_1_1_0_0_1";

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);

            ID = "    <ID>" + scaling_string + "</ID>";
            CVType = "    <Type>1</Type>";
            LimitType = "    <LimitType />";
            LimitMode = "    <LimitMode>0</LimitMode>";
            RangeCorrectionModeForEquiSpacedBP = "    <RangeCorrectionModeForEquiSpacedBP>0</RangeCorrectionModeForEquiSpacedBP>";
            EquiSpacedBPCount = "    <EquiSpacedBPCount>0</EquiSpacedBPCount>";
            G1 = "    <G1>" + Convert.ToString(P.parKa) + "</G1>";
            O1 = "    <O1>" + Convert.ToString(P.parKb) + "</O1>";
            G2 = "    <G2>" + Convert.ToString(P.parKc) + "</G2>";
            O2 = "    <O2>" + Convert.ToString(P.parKd) + "</O2>";

            return (0);
        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <CalibrationScalings>");
            fileXML.WriteLine(ID);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(LimitType);
            fileXML.WriteLine(LimitMode);
            fileXML.WriteLine(RangeCorrectionModeForEquiSpacedBP);
            fileXML.WriteLine(EquiSpacedBPCount);
            fileXML.WriteLine(G1);
            fileXML.WriteLine(O1);
            fileXML.WriteLine(G2);
            fileXML.WriteLine(O2);
            fileXML.WriteLine("  </CalibrationScalings>");

            return (0);
        }

    }

    public class XLSECTSignal
    {
        internal string sigName;
        internal string sigSource;
        internal string sigType;
        internal int sigDecNum;
        internal string sigUnit;
        internal string sigMin;
        internal string sigMax;
        internal string sigKa;
        internal string sigKb;
        internal string sigKc;
        internal string sigKd;
        internal string sigKk;
        internal string sigAlias;
        internal string sigDescription;
        internal string sigBusType;
        internal string sigDimension;
        internal string sigDimA2L;
        internal string sigEvent;
        internal int sigDim;
        internal string sigStatesRow;

        public XLSECTSignal()
        {
            sigName = "";
            sigSource = "";
            sigType = "";
            sigDecNum = 0;
            sigUnit = "";
            sigMin = "";
            sigMax = "";
            sigKa = "";
            sigKb = "";
            sigKc = "";
            sigKd = "";
            sigKk = "";
            sigAlias = "";
            sigDescription = "";
            sigBusType = "";
            sigDimension = "";
            sigDimA2L = "";
            sigEvent = "";
            sigDim = 1;
            sigStatesRow = "";

        }
        public int upload(ref Excel.Worksheet ws, string Line)
        {
            if (ws.get_Range("A" + Line, "A" + Line).Value2 != null)
            {
                /* Common to all types of signals */
                sigName = ws.get_Range("A" + Line, "A" + Line).Value2.ToString();
                sigSource = ws.get_Range("B" + Line, "B" + Line).Value2.ToString();
                sigType = ws.get_Range("C" + Line, "C" + Line).Value2.ToString();
                sigDecNum = Convert.ToInt32(ws.get_Range("D" + Line, "D" + Line).Value2.ToString());
                sigUnit = ws.get_Range("E" + Line, "E" + Line).Value2.ToString();
                sigMin = ws.get_Range("F" + Line, "F" + Line).Value2.ToString();
                sigMax = ws.get_Range("G" + Line, "G" + Line).Value2.ToString();
                sigKa = ws.get_Range("H" + Line, "H" + Line).Value2.ToString();
                sigKb = ws.get_Range("I" + Line, "I" + Line).Value2.ToString();
                sigKc = ws.get_Range("J" + Line, "J" + Line).Value2.ToString();
                sigKd = ws.get_Range("K" + Line, "K" + Line).Value2.ToString();
                sigKk = ws.get_Range("L" + Line, "L" + Line).Value2.ToString();
                sigAlias = ws.get_Range("M" + Line, "M" + Line).Value2.ToString();
                sigDescription = ws.get_Range("N" + Line, "N" + Line).Value2.ToString();
                sigBusType = ws.get_Range("O" + Line, "O" + Line).Value2.ToString();
                sigDimension = ws.get_Range("P" + Line, "P" + Line).Value2.ToString();
                sigDimA2L = ws.get_Range("Q" + Line, "Q" + Line).Value2.ToString();
                if (ws.get_Range("R" + Line, "R" + Line).Value2 != null)
                { 
                    sigEvent = ws.get_Range("R" + Line, "R" + Line).Value2.ToString();
                }
                sigStatesRow = ws.get_Range("S" + Line, "S" + Line).Value2.ToString();

                sigDim = Convert.ToInt32(sigDimA2L);

                if (sigBusType == "O")
                {
                    if (sigDim < 2)
                    {
                        return (0);
                    }
                    else
                    {
                        return (sigDim);
                    }
                }
                else
                {
                    return (-1);
                }

            }
            else
            {
                return (-2); // End of file
            }
        }
    }

    public class SignalValue
    {
        /*
          <Channels>
            <TabAdr>284</TabAdr>
            <FactoryName>CAN_CHA_b.buffer_28[2].b[6]</FactoryName>
            <ASAPName>CAN_CHA_b.buffer_28[2].b[6]</ASAPName>
            <NByte>1</NByte>
            <Format>###0</Format>
            <Unit />
            <Type>512</Type>
            <Default>0</Default>
            <VarNameL0>CAN_CHA_b.buffer_28[2].b[6]</VarNameL0>
            <VarNameL1>CAN_CHA_b.buffer_28[2].b[6]</VarNameL1>
            <DescL0>CAN_CHA_b.buffer_28[2].b[6]</DescL0>
            <DescL1>CAN_CHA_b.buffer_28[2].b[6]</DescL1>
            <Address>0x40001F0E</Address>
            <MinGraph>0</MinGraph>
            <MaxGraph>0</MaxGraph>
            <CorrectionType>0</CorrectionType>
            <LoggerMaxPer>0</LoggerMaxPer>
            <LoggerNBytes>1</LoggerNBytes>
            <Logger>false</Logger>
            <Signed>false</Signed>
            <NByteSingleValue>1</NByteSingleValue>
            <Exportable>true</Exportable>
            <Channel_ScalingID>S_1_1_0_0_1</Channel_ScalingID>
            <Min>0</Min>
            <Max>0</Max>
            <Validated>true</Validated>
            <GroupID />
            <IsArray>true</IsArray>
            <UseMaxSize>false</UseMaxSize>
            <Notes />
            <Open>false</Open>
            <ELFVarType>1</ELFVarType>
          </Channels>
         */
        internal string TabAdr;
        internal string FactoryName;
        internal string ASAPName;
        internal string NByte;
        internal string Format;
        internal string Unit;
        internal string CVType;
        internal string Default;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;
        internal string Address;
        internal string MinGraph;
        internal string MaxGraph;
        internal string CorrectionType;
        internal string LoggerMaxPer;
        internal string LoggerNBytes;
        internal string Logger;
        internal string Signed;
        internal string NByteSingleValue;
        internal string Exportable;
        internal string Channel_ScalingID;
        internal string Min;
        internal string Max;
        internal string Validated;
        internal string GroupID;
        internal string IsArray;
        internal string UseMaxSize;
        internal string Notes;
        internal string Open;
        internal string ELFVarType;

        /* index */

        static int tab_adr = 0;
        public SignalValue()
        {
            TabAdr = "    <TabAdr>0</TabAdr>";
            FactoryName = "    <FactoryName>ZZZZZ</FactoryName>";
            ASAPName = "    <ASAPName></ASAPName>";
            NByte = "    <NByte>2</NByte>";
            Format = "    <Format>####0</Format>";
            Unit = "    <Unit></Unit>";
            CVType = "    <Type>1</Type>";
            Default = "    <Default>0</Default>";
            VarNameL0 = "    <VarNameL0>ZZZZZ</VarNameL0>";
            VarNameL1 = "    <VarNameL1>ZZZZZ</VarNameL1>";
            DescL0 = "    <DescL0></DescL0>";
            DescL1 = "    <DescL1></DescL1>";
            Address = "    <Address>0x00000000</Address>";
            MinGraph = "     <MinGraph>0</MinGraph>";
            MaxGraph = "    <MaxGraph>0</MaxGraph>";
            CorrectionType = "    <CorrectionType>0</CorrectionType>";
            LoggerMaxPer = "    <LoggerMaxPer>0</LoggerMaxPer>";
            LoggerNBytes = "    <LoggerNBytes>1</LoggerNBytes>";
            Logger = "    <Logger>false</Logger>";
            Signed = "    <Signed>false</Signed>";
            NByteSingleValue = "    <NByteSingleValue>2</NByteSingleValue>";
            Exportable = "    <Exportable>true</Exportable>";
            Channel_ScalingID = "    <Channel_ScalingID>S_1_1_0_0_1</Channel_ScalingID>>";
            Min = "    <Var_Min></Var_Min>";
            Max = "    <Var_Max></Var_Max>";
            Validated = "    <Validated></Validated>";
            GroupID = "    <GroupID></GroupID>";
            IsArray = "    <IsArray>false</IsArray>";
            UseMaxSize = "    <UseMaxSize>false</UseMaxSize>";
            Notes = "    <Notes />";
            Open = "    <Open>false</Open>";
            ELFVarType = "    <ELFVarType>1</ELFVarType>";
        }

        public int upload(ref XLSECTSignal P, bool isAnArray, int ind, ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "#######0";
            string type_string = "0";
            string _name = null;
            string i_hex;

            scaling_string = "S_1_" + P.sigKa + "_" + P.sigKb + "_" + P.sigKc + "_" + P.sigKd;

            if (P.sigDecNum != 0)
            {
                format_string += ".";
                for (i = 0; i < P.sigDecNum; i++)
                {
                    format_string += "0";
                }
            }
            switch (P.sigType)
            {
                case "UBYTE":
                    byte_dim = "1";
                    signed_string = "false";
                    type_string = "0";
                    break;
                case "SBYTE":
                    byte_dim = "1";
                    signed_string = "true";
                    type_string = "256";
                    break;
                case "UWORD":
                    byte_dim = "2";
                    signed_string = "false";
                    type_string = "0";
                    break;
                case "SWORD":
                    byte_dim = "2";
                    signed_string = "true";
                    type_string = "256";
                    break;
                case "ULONG":
                    byte_dim = "4";
                    signed_string = "false";
                    type_string = "0";
                    break;
                case "SLONG":
                    byte_dim = "4";
                    signed_string = "true";
                    type_string = "256";
                    break;
                default:
                    byte_dim = "1";
                    signed_string = "false";
                    break;
            }

            _name = P.sigName;

            TabAdr = "    <TabAdr>" + Convert.ToString(tab_adr++) + "</TabAdr>";
            FactoryName = "    <FactoryName>" + P.sigName + "</FactoryName>";
            ASAPName = "    <ASAPName>" + P.sigName + "</ASAPName>";
            NByte = "    <NByte>" + byte_dim + "</NByte>";
            Format = "    <Format>####0</Format>";
            Unit = "    <Unit>" + P.sigUnit + "</Unit>";
            CVType = "    <Type>"+type_string+"</Type>";
            Default = "    <Default>0</Default>";
            VarNameL0 = "    <VarNameL0>" + P.sigName + "</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.sigAlias + "</VarNameL1>";
            DescL0 = "    <DescL0>" + P.sigDescription + "</DescL0>";
            DescL1 = "    <DescL1>" + P.sigDescription + "</DescL1>";
            Address = "    <Address>0xF1CACAF1</Address>";
            MinGraph = "    <MinGraph>" + P.sigMin + "</MinGraph>";
            MaxGraph = "    <MaxGraph>" + P.sigMax + "</MaxGraph>";
            CorrectionType = "    <CorrectionType>0</CorrectionType>";
            LoggerMaxPer = "    <LoggerMaxPer>0</LoggerMaxPer>";
            LoggerNBytes = "    <LoggerNBytes>1</LoggerNBytes>";
            Logger = "    <Logger>false</Logger>";
            Signed = "    <Signed>"+signed_string+"</Signed>";
            NByteSingleValue = "    <NByteSingleValue>" + byte_dim + "</NByteSingleValue>";
            Exportable = "    <Exportable>true</Exportable>";
            Channel_ScalingID = "    <Channel_ScalingID>"+scaling_string+"</Channel_ScalingID>";
            Min = "    <Min>" + P.sigMin + "</Min>";
            Max = "    <Max>" + P.sigMax + "</Max>";
            Validated = "    <Validated>true</Validated>";
            GroupID = "    <GroupID>" + P.sigSource + "</GroupID>";
            IsArray = "    <IsArray>false</IsArray>";
            UseMaxSize = "    <UseMaxSize>false</UseMaxSize>";
            Notes = "    <Notes />";
            Open = "    <Open>false</Open>";
            ELFVarType = "    <ELFVarType>1</ELFVarType>";

            if (isAnArray)
            { 
                FactoryName = "    <FactoryName>" + P.sigName + "["+Convert.ToString(ind)+"]</FactoryName>";
                ASAPName = "    <ASAPName>" + P.sigName + "[" + Convert.ToString(ind) + "]</ASAPName>";
                VarNameL0 = "    <VarNameL0>" + P.sigName + "[" + Convert.ToString(ind) + "]</VarNameL0>";
                VarNameL1 = "    <VarNameL1>" + P.sigAlias + "[" + Convert.ToString(ind) + "]</VarNameL1>";
                IsArray = "    <IsArray>true</IsArray>";
//                ELFVarType = "    <ELFVarType>2</ELFVarType>";
            }

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolBssList.Find(x => x.name == _name);
                    if (isAnArray)
                    {
                        i = Convert.ToInt32(S.address, 16) + ind* Convert.ToInt32(byte_dim);
                        i_hex = i.ToString("X");
                        Address = "    <Address>0x" + i_hex + "</Address>";
                    }
                    else
                    {
                        Address = "    <Address>0x" + S.address + "</Address>";
                    }
                }
                catch
                {
                    MessageBox.Show(_name + "doe not exist");
                }
//                MessageBox.Show(_name + S.name + S.address);
            }
            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <Channels>");
            fileXML.WriteLine(TabAdr);
            fileXML.WriteLine(FactoryName);
            fileXML.WriteLine(ASAPName);
            fileXML.WriteLine(NByte);
            fileXML.WriteLine(Format);
            fileXML.WriteLine(Unit);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(Default);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);
            fileXML.WriteLine(Address);
            fileXML.WriteLine(MinGraph);
            fileXML.WriteLine(MaxGraph);
            fileXML.WriteLine(CorrectionType);
            fileXML.WriteLine(LoggerMaxPer);
            fileXML.WriteLine(LoggerNBytes);
            fileXML.WriteLine(Logger);
            fileXML.WriteLine(Signed);
            fileXML.WriteLine(NByteSingleValue);
            fileXML.WriteLine(Exportable);
            fileXML.WriteLine(Channel_ScalingID);
            fileXML.WriteLine(Min);
            fileXML.WriteLine(Max);
            fileXML.WriteLine(Validated);
            fileXML.WriteLine(GroupID);
            fileXML.WriteLine(IsArray);
            fileXML.WriteLine(UseMaxSize);
            fileXML.WriteLine(Notes);
            fileXML.WriteLine(Open);
            fileXML.WriteLine(ELFVarType);
            fileXML.WriteLine("  </Channels>");

            return (0);
        }

    }

}
