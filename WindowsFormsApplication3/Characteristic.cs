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

    public class XLSECTParameter
    {
        internal string parName;
        internal char parCalType;
        internal string parSource;
        internal string parType;
        internal int parDecNum;
        internal string parUnit;
        internal string parMin;
        internal string parMax;
        internal int parKa;
        internal int parKb;
        internal int parKc;
        internal int parKd;
        internal int parKk;
        internal string parAlias;
        internal string parDescription;
        internal string parDimension;
        internal string parBreakpoint1;
        internal string parBreakpoint2;
        internal string parInputQuantity;
        internal string parStatesRow;
        internal int parDim1;
        internal int parDim2;

        public XLSECTParameter()
        {
            parName = "";
            parCalType = 'Z';
            parSource = "";
            parType = "";
            parDecNum = 0;
            parUnit = "";
            parMin = "";
            parMax = "";
            parKa = 1;
            parKb = 0;
            parKc = 0;
            parKd = 1;
            parKk = 1;
            parAlias = "";
            parDescription = "";
            parDimension = "";
            parBreakpoint1 = "";
            parBreakpoint2 = "";
            parInputQuantity = "";
            parStatesRow = "";
            parDim1 = 1;
            parDim2 = 1;
        }

        public int upload(ref Excel.Worksheet ws, string Line)
        {
            int index;
            string Dim1;
            string Dim2;
            char x;


            if (ws.get_Range("A" + Line, "A" + Line).Value2 != null)
            {
                /* Common to all types of paramteters */
                parName = ws.get_Range("A" + Line, "A" + Line).Value2.ToString();
                parCalType = parName[parName.Length - 1];
                if (parCalType == 'k') parCalType = 'B';
                parSource = ws.get_Range("B" + Line, "B" + Line).Value2.ToString();
                parType = ws.get_Range("C" + Line, "C" + Line).Value2.ToString();
                parDecNum = Convert.ToInt32(ws.get_Range("D" + Line, "D" + Line).Value2.ToString());
                parUnit = ws.get_Range("E" + Line, "E" + Line).Value2.ToString();
                parMin = ws.get_Range("F" + Line, "F" + Line).Value2.ToString();
                parMax = ws.get_Range("G" + Line, "G" + Line).Value2.ToString();
                parKa = Convert.ToInt32(ws.get_Range("H" + Line, "H" + Line).Value2.ToString());
                parKb = Convert.ToInt32(ws.get_Range("I" + Line, "I" + Line).Value2.ToString());
                parKc = Convert.ToInt32(ws.get_Range("J" + Line, "J" + Line).Value2.ToString());
                parKd = Convert.ToInt32(ws.get_Range("K" + Line, "K" + Line).Value2.ToString());
                parKk = Convert.ToInt32(ws.get_Range("L" + Line, "L" + Line).Value2.ToString());
                parAlias = ws.get_Range("M" + Line, "M" + Line).Value2.ToString();
                parDescription = ws.get_Range("N" + Line, "N" + Line).Value2.ToString();
                if (ws.get_Range("O" + Line, "O" + Line).Value2 != null)
                { 
                    parDimension = ws.get_Range("O" + Line, "O" + Line).Value2.ToString();
                }
                else
                {
                    parDimension = "1x1";
                }
                parStatesRow = ws.get_Range("T" + Line, "T" + Line).Value2.ToString();

//                DimDim = parDimension.Length;

                index = 0;
                x = parDimension[index];
                Dim1 = "";
                Dim2 = "";

                while(x != 'x')
                {   
                    Dim1 += parDimension[index++];
                    x = parDimension[index];
                }
                index++;
                while (index < parDimension.Length)
                {
                    Dim2 += parDimension[index++];
                }

                parDim1 = Convert.ToInt32(Dim1);
                parDim2 = Convert.ToInt32(Dim2);

                if (parDimension == "1x1")
                {

                    if (ws.get_Range("P" + Line, "P" + Line).Value2 != null)
                    {
                        parBreakpoint1 = ws.get_Range("P" + Line, "P" + Line).Value2.ToString();
                    }
                    if (ws.get_Range("Q" + Line, "Q" + Line).Value2 != null)
                    {
                        parBreakpoint2 = ws.get_Range("Q" + Line, "Q" + Line).Value2.ToString();
                    }
                    if (ws.get_Range("S" + Line, "S" + Line).Value2 != null)
                    {
                        parInputQuantity = ws.get_Range("S" + Line, "S" + Line).Value2.ToString();
                    }
//                    MessageBox.Show(parName);

                    return (0); // CalibrationValue = 0

                }
                else
                {
                    /* Shared axis or map/curve without axes */
                    if (ws.get_Range("P" + Line, "P" + Line).Value2 == null)
                    {
                        if (ws.get_Range("S" + Line, "S" + Line).Value2 != null)   /* Shared axis */
                        {
                            if (parCalType == 'B')
                            { 
                                parInputQuantity = ws.get_Range("S" + Line, "S" + Line).Value2.ToString();
                                return (1); // Shared axis = 1
                            }
                            else
                            {
                                return (-1); /* Discrepancy beetween parCalType and no parInputQuantity */
                            }

                        }
                        else
                        {
                            /* VAL_BLK array or matrix */
                            if (parDim1 == 1)
                            {
                                if ((ws.get_Range("S" + Line, "S" + Line).Value2 == "NONE")
                                    || (ws.get_Range("S" + Line, "S" + Line).Value2 == "none")
                                    || (ws.get_Range("S" + Line, "S" + Line).Value2 == "None")
                                   )
                                {
                                    if ((parCalType == 'V') || (parCalType == 'C'))
                                    { 
                                        return (4);  /* VAL_BLK array */
                                    }
                                    else
                                    {
                                        return (-1); /* Discrepancy beetween parCalType and VAL_BLK array without axes */
                                    }
                                }
                                else
                                {
                                    if (parCalType == 'V')
                                    {
                                        return (5);  /* CURVE with embedded fixed axis */
                                    }
                                    else
                                    {
                                        return (-1); /* Discrepancy beetween parCalType and CURVE with embedded axes */
                                    }
                                }
                            }
                            else
                            {
                                if ((ws.get_Range("S" + Line, "S" + Line).Value2 == "NONE")
                                    || (ws.get_Range("S" + Line, "S" + Line).Value2 == "none")
                                    || (ws.get_Range("S" + Line, "S" + Line).Value2 == "None")
                                   )
                                {
                                    if ((parCalType == 'T') || (parCalType == 'C'))
                                    {
                                        return (6);  /* VAL_BLK 2 dim matrix */
                                    }
                                    else
                                    {
                                        return (-1); /* Discrepancy beetween parCalType and VAL_BLK matrix without axes */
                                    }
                                }
                                else
                                {
                                    if (parCalType == 'T')
                                    {
                                        return (7);  /* MAP with embedded fixed axes */
                                    }
                                    else
                                    {
                                        return (-1); /* Discrepancy beetween parCalType and MAP with embedded fixed axes */
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        parBreakpoint1 = ws.get_Range("P" + Line, "P" + Line).Value2.ToString();

                        /* There is a breakpoint 2 => it's a 2-DIM table */
                        if (ws.get_Range("Q" + Line, "Q" + Line).Value2 != null)
                        {
                            parBreakpoint2 = ws.get_Range("Q" + Line, "Q" + Line).Value2.ToString();
                            return (3); // 2-DIM Interpolation map with shared axes = 3
                        }
                        else
                        {
                            /* There is no breakpoint 2 => it's a 1-DIM table */
                            return (2); // 1-DIM Interpolation map with shared axes = 2

                        }                    
                    }
                }
            }
            else
            {
                return (-2);
            }
        }
    }
    public class CalibrationValue
    {
        /*
            <CalibrationValues>
            <FactoryName>EngStrt_rpmthresEngineOn_C</FactoryName>
            <CalibrationTypeID>LIE00</CalibrationTypeID>
            <Address>0x00A00A08</Address>
            <OffsetByte>2568</OffsetByte>
            <OffsetBit>0</OffsetBit>
            <GroupID>ENGST</GroupID>
            <Type>1</Type>
            <NByte>2</NByte>
            <NByteSingleValue>2</NByteSingleValue>
            <NBit>0</NBit>
            <Signed>false</Signed>
            <VarNameL0>EngStrt_rpmthresEngineOn_C</VarNameL0>
            <VarNameL1>EngStrt_rpmthresEngineOn_C</VarNameL1>
            <DescL0>rpm threshold above which Engine is On</DescL0>
            <DescL1>rpm threshold above which Engine is On</DescL1>
            <Var_ScalingID>S_1_1_0_0_1</Var_ScalingID>
            <Var_Format>####0</Var_Format>
            <Var_Unit>Rpm</Var_Unit>
            <Var_Min>0</Var_Min>
            <Var_Max>65535</Var_Max>
            <Var_MinEdit>0</Var_MinEdit>
            <Var_MaxEdit>65535</Var_MaxEdit>
            <Exportable>true</Exportable>
            <Validated>true</Validated>
            <SubType>0</SubType>
            <IsArray>false</IsArray>
            <UseMaxSize>false</UseMaxSize>
            <Open>false</Open>
            <Notes />
            <ELFVarType>1</ELFVarType>
            <Var_FormulaID />
          </CalibrationValues>
        */
        internal string FactoryName;
        internal string CalibrationTypeID;
        internal string Address;
        internal string OffsetByte;
        internal string OffsetBit;
        internal string GroupID;
        internal string CVType;
        internal string NByte;
        internal string NByteSingleValue;
        internal string NBit;
        internal string Signed;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;
        internal string Var_ScalingID;
        internal string Var_Format;
        internal string Var_Unit;
        internal string Var_Min;
        internal string Var_Max;
        internal string Var_MinEdit;
        internal string Var_MaxEdit;
        internal string Exportable;
        internal string Validated;
        internal string SubType;
        internal string IsArray;
        internal string UseMaxSize;
        internal string Open;
        internal string Notes;
        internal string ELFVarType;
        internal string Var_FormulaID;


        public int upload(ref XLSECTParameter P,ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "#######0";
            string _name = null;


            scaling_string = "S_1_"+Convert.ToString(P.parKa)+"_"+Convert.ToString(P.parKb)+"_"+Convert.ToString(P.parKc)+"_"+Convert.ToString(P.parKd);
            if (P.parDecNum != 0)
            {
                format_string += ".";
                for (i = 0; i < P.parDecNum; i++)
                {
                    format_string += "0";
                }
            }
            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = "1";
                    signed_string = "false";
                    break;
                case "SBYTE":
                    byte_dim = "1";
                    signed_string = "true";
                    break;
                case "UWORD":
                    byte_dim = "2";
                    signed_string = "false";
                    break;
                case "SWORD":
                    byte_dim = "2";
                    signed_string = "true";
                    break;
                case "ULONG":
                    byte_dim = "4";
                    signed_string = "false";
                    break;
                case "SLONG":
                    byte_dim = "4";
                    signed_string = "true";
                    break;
                default:
                    byte_dim = "1";
                    signed_string = "false";
                    break;
            }
            _name = P.parName;

            FactoryName = "    <FactoryName>"+P.parName+"</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>" + P.parSource + "</GroupID>";
            CVType = "    <Type>1</Type>";
            NByte = "    <NByte>" + byte_dim + "</NByte>";
            NByteSingleValue = "    <NByteSingleValue>" + byte_dim + "</NByteSingleValue>";
            NBit = "    <NBit>0</NBit>";
            Signed = "    <Signed>" + signed_string + "</Signed>";
            VarNameL0 = "    <VarNameL0>" + P.parAlias + "</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.parName + "</VarNameL1>";
            DescL0 = "    <DescL0>" + P.parDescription + "</DescL0>";
            DescL1 = "    <DescL1>"+P.parDescription+"</DescL1>";
            Var_ScalingID = "    <Var_ScalingID>"+scaling_string+"</Var_ScalingID>";
            Var_Format = "    <Var_Format>"+format_string+"</Var_Format>";
            Var_Unit = "    <Var_Unit>"+P.parUnit+"</Var_Unit>";
            Var_Min = "    <Var_Min>"+P.parMin+"</Var_Min>";
            Var_Max = "    <Var_Max>"+P.parMax+"</Var_Max>";
            Var_MinEdit = "    <Var_MinEdit>"+Convert.ToString(P.parMin)+"</Var_MinEdit>";
            Var_MaxEdit = "    <Var_MaxEdit>"+Convert.ToString(P.parMax)+"</Var_MaxEdit>";
            Exportable = "    <Exportable>true</Exportable>";
            Validated = "    <Validated>true</Validated>";
            SubType = "    <SubType>0</SubType>";
            IsArray = "    <IsArray>false</IsArray>";
            UseMaxSize = "    <UseMaxSize>false</UseMaxSize>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            ELFVarType = "    <ELFVarType>1</ELFVarType>";
            Var_FormulaID = "    <Var_FormulaID />";

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
                    Address = "    <Address>0x" + S.address + "</Address>";
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the symbol table");
                }
                //                MessageBox.Show(_name + S.name + S.address);
            }

            return (0);

        }

        public CalibrationValue()
        {
            FactoryName = "    <FactoryName>ZZZZZ</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>ENGST</GroupID>";
            CVType = "    <Type>1</Type>";
            NByte = "    <NByte>2</NByte>";
            NByteSingleValue = "    <NByteSingleValue>2</NByteSingleValue>";
            NBit = "    <NBit>0</NBit>";
            Signed = "    <Signed>false</Signed>";
            VarNameL0 = "    <VarNameL0>ZZZZZ</VarNameL0>";
            VarNameL1 = "    <VarNameL1>ZZZZZ</VarNameL1>";
            DescL0 = "    <DescL0>rpm threshold above which Engine is On</DescL0>";
            DescL1 = "    <DescL1>rpm threshold above which Engine is On</DescL1>";
            Var_ScalingID = "    <Var_ScalingID>S_1_1_0_0_1</Var_ScalingID>";
            Var_Format = "    <Var_Format>####0</Var_Format>";
            Var_Unit = "    <Var_Unit></Var_Unit>";
            Var_Min = "    <Var_Min>0</Var_Min>";
            Var_Max = "    <Var_Max>65535</Var_Max>";
            Var_MinEdit = "    <Var_MinEdit>0</Var_MinEdit>";
            Var_MaxEdit = "    <Var_MaxEdit>65535</Var_MaxEdit>";
            Exportable = "    <Exportable>true</Exportable>";
            Validated = "    <Validated>true</Validated>";
            SubType = "    <SubType>0</SubType>";
            IsArray = "    <IsArray>false</IsArray>";
            UseMaxSize = "    <UseMaxSize>false</UseMaxSize>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            ELFVarType = "    <ELFVarType>1</ELFVarType>";
            Var_FormulaID = "    <Var_FormulaID />";
        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <CalibrationValues>");
            fileXML.WriteLine(FactoryName);
            fileXML.WriteLine(CalibrationTypeID);
            fileXML.WriteLine(Address);
            fileXML.WriteLine(OffsetByte);
            fileXML.WriteLine(OffsetBit);
            fileXML.WriteLine(GroupID);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(NByte);
            fileXML.WriteLine(NByteSingleValue);
            fileXML.WriteLine(NBit);
            fileXML.WriteLine(Signed);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);
            fileXML.WriteLine(Var_ScalingID);
            fileXML.WriteLine(Var_Format);
            fileXML.WriteLine(Var_Unit);
            fileXML.WriteLine(Var_Min);
            fileXML.WriteLine(Var_Max);
            fileXML.WriteLine(Var_MinEdit);
            fileXML.WriteLine(Var_MaxEdit);
            fileXML.WriteLine(Exportable);
            fileXML.WriteLine(Validated);
            fileXML.WriteLine(SubType);
            fileXML.WriteLine(IsArray);
            fileXML.WriteLine(UseMaxSize);
            fileXML.WriteLine(Open);
            fileXML.WriteLine(Notes);
            fileXML.WriteLine(ELFVarType);
            fileXML.WriteLine(Var_FormulaID);
            fileXML.WriteLine("  </CalibrationValues>");

            return (0);
        }

        public void Show ()
        {
            MessageBox.Show(
                            "  <CalibrationValues>" + "\r\n" +
                                  FactoryName + "\r\n" +
                                  CalibrationTypeID+ "\r\n" +
                                  Address+ "\r\n" +
                                  OffsetByte+ "\r\n" +
                                  OffsetBit+ "\r\n" +
                                  GroupID+ "\r\n" +
                                  CVType+ "\r\n" +
                                  NByte+ "\r\n" +
                                  NByteSingleValue+ "\r\n" +
                                  NBit+ "\r\n" +
                                  Signed+ "\r\n" +
                                  VarNameL0+ "\r\n" +
                                  VarNameL1+ "\r\n" +
                                  DescL0+ "\r\n" +
                                  DescL1+ "\r\n" +
                                  Var_ScalingID+ "\r\n" +
                                  Var_Format+ "\r\n" +
                                  Var_Unit+ "\r\n" +
                                  Var_Min+ "\r\n" +
                                  Var_Max+ "\r\n" +
                                  Var_MinEdit+ "\r\n" +
                                  Var_MaxEdit+ "\r\n" +
                                  Exportable+ "\r\n" +
                                  Validated+ "\r\n" +
                                  SubType+ "\r\n" +
                                  IsArray+ "\r\n" +
                                  UseMaxSize+ "\r\n" +
                                  Open+ "\r\n" +
                                  Notes+ "\r\n" +
                                  ELFVarType+ "\r\n" +
                                  Var_FormulaID+ "\r\n" + "  </CalibrationValues>" 
               );

        }
    }
    public class CalibrationCurve
    {
        /*
           <CalibrationCurves>
            <FactoryName>FuelPL_fPIIntegrTerm_V[0]</FactoryName>
            <CalibrationTypeID>LIE00</CalibrationTypeID>
            <Address>0x00A09BD0</Address>
            <OffsetByte>39888</OffsetByte>
            <OffsetBit>0</OffsetBit>
            <GroupID>FUELPL</GroupID>
            <Type>64</Type>
            <NByte>16</NByte>
            <NByteSingleValue>2</NByteSingleValue>
            <EquiSpaced>false</EquiSpaced>
            <Signed>true</Signed>
            <VarNameL0>FuelPL_fPIIntegrTerm_V[0]</VarNameL0>
            <VarNameL1>FuelPL_fPIIntegrTerm_V[0]</VarNameL1>
            <DescL0>FuelPL_fPIIntegrTerm_V[0]</DescL0>
            <DescL1>FuelPL_fPIIntegrTerm_V[0]</DescL1>
            <Var_Label />
            <Var_ScalingID>S_1_1_0_0_1</Var_ScalingID>
            <Var_Format>####0</Var_Format>
            <Var_Unit />
            <Var_Min>-32768</Var_Min>
            <Var_Max>32767</Var_Max>
            <Var_MinEdit>0</Var_MinEdit>
            <Var_MaxEdit>0</Var_MaxEdit>
            <Var_ReferenceChannel />
            <BreakPoint1_Label />
            <BreakPoint1_ScalingID />
            <BreakPoint1_Format />
            <BreakPoint1_Unit />
            <BreakPoint1_Min>0</BreakPoint1_Min>
            <BreakPoint1_Max>0</BreakPoint1_Max>
            <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>
            <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>
            <BreakPoint1_ReferenceChannel />
            <BreakPoint1_Count>0</BreakPoint1_Count>
            <Exportable>true</Exportable>
            <Validated>true</Validated>
            <BreakPoint1_FactoryName>FuelPL_fPITerm_Bk[0]</BreakPoint1_FactoryName>
            <IsArray>true</IsArray>
            <UseMaxSize>true</UseMaxSize>
            <Open>false</Open>
            <Notes />
            <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>
            <ELFVarType>2</ELFVarType>
            <Var_IncrementID />
            <Var_FormulaID />
            <BreakPoint1_FormulaID />
            </CalibrationCurves>
 
        */
        internal string FactoryName;
        internal string CalibrationTypeID;
        internal string Address;
        internal string OffsetByte;
        internal string OffsetBit;
        internal string GroupID;
        internal string CVType;
        internal string NByte;
        internal string NByteSingleValue;
        internal string Equispaced;
        internal string Signed;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;
        internal string Var_Label;
        internal string Var_ScalingID;
        internal string Var_Format;
        internal string Var_Unit;
        internal string Var_Min;
        internal string Var_Max;
        internal string Var_MinEdit;
        internal string Var_MaxEdit;
        internal string Var_ReferenceChannel;
        internal string BreakPoint1_Label;
        internal string BreakPoint1_ScalingID;
        internal string BreakPoint1_Format;
        internal string BreakPoint1_Unit;
        internal string BreakPoint1_Min;
        internal string BreakPoint1_Max;
        internal string BreakPoint1_MinEdit;
        internal string BreakPoint1_MaxEdit;
        internal string BreakPoint1_ReferenceChannel;
        internal string BreakPoint1_Count;
        internal string Exportable;
        internal string Validated;
        internal string BreakPoint1_FactoryName;
        internal string IsArray;
        internal string UseMaxSize;
        internal string Open;
        internal string Notes;
        internal string BreakPoint1_Monotonicity;
        internal string ELFVarType;
        internal string Var_IncrementID;
        internal string Var_FormulaID;
        internal string BreakPoint1_FormulaID;


        public CalibrationCurve()
        {
            FactoryName = "    <FactoryName>ZZZZZ</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>XXXXX</GroupID>";
            CVType = "    <Type>64</Type>";
            NByte = "    <NByte>2</NByte>";
            NByteSingleValue = "    <NByteSingleValue>2</NByteSingleValue>";
            Equispaced = "    <EquiSpaced>false</EquiSpaced>";
            Signed = "    <Signed>false</Signed>";
            VarNameL0 = "    <VarNameL0>ZZZZZ</VarNameL0>";
            VarNameL1 = "    <VarNameL1>ZZZZZ</VarNameL1>";
            DescL0 = "    <DescL0>rpm threshold above which Engine is On</DescL0>";
            DescL1 = "    <DescL1>rpm threshold above which Engine is On</DescL1>";
            Var_Label = "    <Var_Label />";
            Var_ScalingID = "    <Var_ScalingID>S_1_1_0_0_1</Var_ScalingID>";
            Var_Format = "    <Var_Format>####0</Var_Format>";
            Var_Unit = "    <Var_Unit/>";
            Var_Min = "    <Var_Min>0</Var_Min>";
            Var_Max = "    <Var_Max>65535</Var_Max>";
            Var_MinEdit = "    <Var_MinEdit>0</Var_MinEdit>";
            Var_MaxEdit = "    <Var_MaxEdit>65535</Var_MaxEdit>";
            Var_ReferenceChannel = "    <Var_ReferenceChannel />";
            BreakPoint1_Label = "    <BreakPoint1_Label />";
            BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID />";
            BreakPoint1_Format = "    <BreakPoint1_Format />";
            BreakPoint1_Unit  = "    <BreakPoint1_Unit />";
            BreakPoint1_Min = "    <BreakPoint1_Min>0</BreakPoint1_Min>";
            BreakPoint1_Max = "    <BreakPoint1_Max>0</BreakPoint1_Max>";
            BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
            BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>";
            BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel />";
            BreakPoint1_Count = "    <BreakPoint1_Count>0</BreakPoint1_Count>";
            Exportable = "    <Exportable>true</Exportable>";
            Validated = "    <Validated>true</Validated>";
            BreakPoint1_FactoryName = "    <BreakPoint1_FactoryName>[0]</BreakPoint1_FactoryName>";
            IsArray = "    <IsArray>true</IsArray>";
            UseMaxSize = "    <UseMaxSize>true</UseMaxSize>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
            ELFVarType = "    <ELFVarType>2</ELFVarType>";
            Var_IncrementID = "    <Var_IncrementID />";
            Var_FormulaID = "    <Var_FormulaID />";
            BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";
        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "#######0";
            string type_string = "0";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);
            if (P.parDecNum != 0)
            {
                format_string += ".";
                for (i = 0; i < P.parDecNum; i++)
                {
                    format_string += "0";
                }
            }
            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "62";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "true";
                    type_string = "62";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2*P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "false";
                    type_string = "64";
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "true";
                    type_string = "64";
                    break;
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "false";
                    type_string = "64";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "true";
                    type_string = "64";
                    break;
                default:
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "64";
                    break;
            }

            _name = P.parName;
            //******************************************************

            FactoryName = "    <FactoryName>" + P.parName + "[0]</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>" + P.parSource + "</GroupID>";
            CVType = "    <Type>"+type_string+"</Type>";
            NByte = "    <NByte>" + byte_dim + "</NByte>";
            NByteSingleValue = "    <NByteSingleValue>" + byte_dim_single_value + "</NByteSingleValue>";
            Equispaced = "    <EquiSpaced>false</EquiSpaced>";
            Signed = "    <Signed>" + signed_string + "</Signed>";
            VarNameL0 = "    <VarNameL0>" + P.parAlias + "[0]</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.parName + "[0]</VarNameL1>";
            DescL0 = "    <DescL0>" + P.parDescription + "</DescL0>";
            DescL1 = "    <DescL1>" + P.parDescription + "</DescL1>";
            Var_Label = "    <Var_Label />";
            Var_ScalingID = "    <Var_ScalingID>" + scaling_string + "</Var_ScalingID>";
            Var_Format = "    <Var_Format>" + format_string + "</Var_Format>";
            Var_Unit = "    <Var_Unit>" + P.parUnit + "</Var_Unit>";
            Var_Min = "    <Var_Min>" + P.parMin + "</Var_Min>";
            Var_Max = "    <Var_Max>" + P.parMax + "</Var_Max>";
            Var_MinEdit = "    <Var_MinEdit>" + P.parMin + "</Var_MinEdit>";
            Var_MaxEdit = "    <Var_MaxEdit>" + P.parMax + "</Var_MaxEdit>";
            Var_ReferenceChannel = "    <Var_ReferenceChannel />";
            BreakPoint1_Label = "    <BreakPoint1_Label />";
            BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID />";
            BreakPoint1_Format = "    <BreakPoint1_Format />";
            BreakPoint1_Unit = "    <BreakPoint1_Unit />";
            BreakPoint1_Min = "    <BreakPoint1_Min>0</BreakPoint1_Min>";
            BreakPoint1_Max = "    <BreakPoint1_Max>0</BreakPoint1_Max>";
            BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
            BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>";
            BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel />";
            BreakPoint1_Count = "    <BreakPoint1_Count>0</BreakPoint1_Count>";
            Exportable = "    <Exportable>true</Exportable>";
            Validated = "    <Validated>true</Validated>";
            BreakPoint1_FactoryName = "    <BreakPoint1_FactoryName>"+P.parBreakpoint1+"[0]</BreakPoint1_FactoryName>";
            IsArray = "    <IsArray>true</IsArray>";
            UseMaxSize = "    <UseMaxSize>true</UseMaxSize>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
            ELFVarType = "    <ELFVarType>2</ELFVarType>";
            Var_IncrementID = "    <Var_IncrementID />";
            Var_FormulaID = "    <Var_FormulaID />";
            BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";
            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
                    Address = "    <Address>0x" + S.address + "</Address>";
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the Symbol database");
                }
                //                MessageBox.Show(_name + S.name + S.address);
            }

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <CalibrationCurves>");
            fileXML.WriteLine(FactoryName);
            fileXML.WriteLine(CalibrationTypeID);
            fileXML.WriteLine(Address);
            fileXML.WriteLine(OffsetByte);
            fileXML.WriteLine(OffsetBit);
            fileXML.WriteLine(GroupID);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(NByte);
            fileXML.WriteLine(NByteSingleValue);
//            fileXML.WriteLine(NBit);

            fileXML.WriteLine(Equispaced);

            fileXML.WriteLine(Signed);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);

            fileXML.WriteLine(Var_Label);
            fileXML.WriteLine(Var_ScalingID);
            fileXML.WriteLine(Var_Format);
            fileXML.WriteLine(Var_Unit);
            fileXML.WriteLine(Var_Min);
            fileXML.WriteLine(Var_Max);
            fileXML.WriteLine(Var_MinEdit);
            fileXML.WriteLine(Var_MaxEdit);
            fileXML.WriteLine(Var_ReferenceChannel);

            fileXML.WriteLine(BreakPoint1_Label);
            fileXML.WriteLine(BreakPoint1_ScalingID);
            fileXML.WriteLine(BreakPoint1_Format);
            fileXML.WriteLine(BreakPoint1_Unit);
            fileXML.WriteLine(BreakPoint1_Min);
            fileXML.WriteLine(BreakPoint1_Max);
            fileXML.WriteLine(BreakPoint1_MinEdit);
            fileXML.WriteLine(BreakPoint1_MaxEdit);
            fileXML.WriteLine(BreakPoint1_ReferenceChannel);
            fileXML.WriteLine(BreakPoint1_Count);


            fileXML.WriteLine(Exportable);
            fileXML.WriteLine(Validated);
            fileXML.WriteLine(BreakPoint1_FactoryName);
            fileXML.WriteLine(IsArray);
            fileXML.WriteLine(UseMaxSize);
            fileXML.WriteLine(Open);
            fileXML.WriteLine(Notes);
            fileXML.WriteLine(BreakPoint1_Monotonicity);
            fileXML.WriteLine(ELFVarType);
            fileXML.WriteLine(Var_IncrementID);
            fileXML.WriteLine(Var_FormulaID);
            fileXML.WriteLine(BreakPoint1_FormulaID);
            fileXML.WriteLine("  </CalibrationCurves>");

            return (0);
        }

        public void Show()
        {
            MessageBox.Show("Curves");
        }
    }
    public class CalibrationMap
    {
        /*
          <CalibrationMaps>
            <FactoryName>IgnCtl_degSABaseIntOnExhOn_T[0]</FactoryName>
            <CalibrationTypeID>LIE00</CalibrationTypeID>
            <Address>0x00A05818</Address>
            <OffsetByte>22552</OffsetByte>
            <OffsetBit>0</OffsetBit>
            <GroupID>IGNCTL</GroupID>
            <Type>65</Type>
            <NByte>650</NByte>
            <NByteSingleValue>2</NByteSingleValue>
            <EquiSpaced>false</EquiSpaced>
            <Signed>true</Signed>
            <VarNameL0>IgnCtl_degSABaseIntOnExhOn_T[0]</VarNameL0>
            <VarNameL1>IgnCtl_degSABaseIntOnExhOn_T[0]</VarNameL1>
            <DescL0>Spark advance target base Intake VVT On Exhaust VVT On</DescL0>
            <DescL1>IgnCtl_degSAOptIntOffExhOff_T[0]</DescL1>
            <Var_Label>SABase</Var_Label>
            <Var_ScalingID>S_1_1_0_0_16</Var_ScalingID>
            <Var_Format>####0.000</Var_Format>
            <Var_Unit>deg</Var_Unit>
            <Var_Min>-2048</Var_Min>
            <Var_Max>2047.9375</Var_Max>
            <Var_MinEdit>-50</Var_MinEdit>
            <Var_MaxEdit>70</Var_MaxEdit>
            <Var_ReferenceChannel />
            <BreakPoint1_Label />
            <BreakPoint1_ScalingID />
            <BreakPoint1_Format />
            <BreakPoint1_Unit />
            <BreakPoint1_Min>0</BreakPoint1_Min>
            <BreakPoint1_Max>0</BreakPoint1_Max>
            <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>
            <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>
            <BreakPoint1_ReferenceChannel />
            <BreakPoint1_Count>0</BreakPoint1_Count>
            <BreakPoint2_Label />
            <BreakPoint2_ScalingID />
            <BreakPoint2_Format />
            <BreakPoint2_Unit />
            <BreakPoint2_Min>0</BreakPoint2_Min>
            <BreakPoint2_Max>0</BreakPoint2_Max>
            <BreakPoint2_MinEdit>0</BreakPoint2_MinEdit>
            <BreakPoint2_MaxEdit>0</BreakPoint2_MaxEdit>
            <BreakPoint2_ReferenceChannel />
            <BreakPoint2_Count>0</BreakPoint2_Count>
            <Exportable>true</Exportable>
            <Validated>true</Validated>
            <BreakPoint1_FactoryName>IgnCtl_rpmxSATable_Bk[0]</BreakPoint1_FactoryName>
            <BreakPoint2_FactoryName>IgnCtl_rRCActxSATable_Bk[0]</BreakPoint2_FactoryName>
            <IsArray>true</IsArray>
            <UseMaxSize>true</UseMaxSize>
            <Open>false</Open>
            <Notes />
            <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>
            <BreakPoint2_Monotonicity>1</BreakPoint2_Monotonicity>
            <ELFVarType>2</ELFVarType>
            <Var_IncrementID />
            <Var_FormulaID />
            <BreakPoint1_FormulaID />
            <BreakPoint2_FormulaID />
          </CalibrationMaps>     
       
        */
        internal string FactoryName;
        internal string CalibrationTypeID;
        internal string Address;
        internal string OffsetByte;
        internal string OffsetBit;
        internal string GroupID;
        internal string CVType;
        internal string NByte;
        internal string NByteSingleValue;
        internal string Equispaced;
        internal string Signed;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;
        internal string Var_Label;
        internal string Var_ScalingID;
        internal string Var_Format;
        internal string Var_Unit;
        internal string Var_Min;
        internal string Var_Max;
        internal string Var_MinEdit;
        internal string Var_MaxEdit;
        internal string Var_ReferenceChannel;
        internal string BreakPoint1_Label;
        internal string BreakPoint1_ScalingID;
        internal string BreakPoint1_Format;
        internal string BreakPoint1_Unit;
        internal string BreakPoint1_Min;
        internal string BreakPoint1_Max;
        internal string BreakPoint1_MinEdit;
        internal string BreakPoint1_MaxEdit;
        internal string BreakPoint1_ReferenceChannel;
        internal string BreakPoint1_Count;
        internal string BreakPoint2_Label;
        internal string BreakPoint2_ScalingID;
        internal string BreakPoint2_Format;
        internal string BreakPoint2_Unit;
        internal string BreakPoint2_Min;
        internal string BreakPoint2_Max;
        internal string BreakPoint2_MinEdit;
        internal string BreakPoint2_MaxEdit;
        internal string BreakPoint2_ReferenceChannel;
        internal string BreakPoint2_Count;
        internal string Exportable;
        internal string Validated;
        internal string BreakPoint1_FactoryName;
        internal string BreakPoint2_FactoryName;
        internal string IsArray;
        internal string UseMaxSize;
        internal string Open;
        internal string Notes;
        internal string BreakPoint1_Monotonicity;
        internal string BreakPoint2_Monotonicity;
        internal string ELFVarType;
        internal string Var_IncrementID;
        internal string Var_FormulaID;
        internal string BreakPoint1_FormulaID;
        internal string BreakPoint2_FormulaID;

        public CalibrationMap()
        {
                FactoryName = "    <FactoryName>ZZZZZ</FactoryName>";
                CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
                Address = "    <Address>0x00000000</Address>";
                OffsetByte = "    <OffsetByte>0</OffsetByte>";
                OffsetBit = "    <OffsetBit>0</OffsetBit>";
                GroupID = "    <GroupID>XXXXX</GroupID>";
                CVType = "    <Type>65</Type>";
                NByte = "    <NByte>2</NByte>";
                NByteSingleValue = "    <NByteSingleValue>2</NByteSingleValue>";
                Equispaced = "    <EquiSpaced>false</EquiSpaced>";
                Signed = "    <Signed>false</Signed>";
                VarNameL0 = "    <VarNameL0>ZZZZZ</VarNameL0>";
                VarNameL1 = "    <VarNameL1>ZZZZZ</VarNameL1>";
                DescL0 = "    <DescL0>rpm threshold above which Engine is On</DescL0>";
                DescL1 = "    <DescL1>rpm threshold above which Engine is On</DescL1>";
                Var_Label = "    <Var_Label />";
                Var_ScalingID = "    <Var_ScalingID>S_1_1_0_0_1</Var_ScalingID>";
                Var_Format = "    <Var_Format>####0</Var_Format>";
                Var_Unit = "    <Var_Unit/>";
                Var_Min = "    <Var_Min>0</Var_Min>";
                Var_Max = "    <Var_Max>65535</Var_Max>";
                Var_MinEdit = "    <Var_MinEdit>0</Var_MinEdit>";
                Var_MaxEdit = "    <Var_MaxEdit>65535</Var_MaxEdit>";
                Var_ReferenceChannel = "    <Var_ReferenceChannel />";
                BreakPoint1_Label = "    <BreakPoint1_Label />";
                BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID />";
                BreakPoint1_Format = "    <BreakPoint1_Format />";
                BreakPoint1_Unit = "    <BreakPoint1_Unit />";
                BreakPoint1_Min = "    <BreakPoint1_Min>0</BreakPoint1_Min>";
                BreakPoint1_Max = "    <BreakPoint1_Max>0</BreakPoint1_Max>";
                BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
                BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>";
                BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel />";
                BreakPoint1_Count = "    <BreakPoint1_Count>0</BreakPoint1_Count>";
                BreakPoint2_Label = "    <BreakPoint2_Label />";
                BreakPoint2_ScalingID = "    <BreakPoint2_ScalingID />";
                BreakPoint2_Format = "    <BreakPoint2_Format />";
                BreakPoint2_Unit = "    <BreakPoint2_Unit />";
                BreakPoint2_Min = "    <BreakPoint2_Min>0</BreakPoint2_Min>";
                BreakPoint2_Max = "    <BreakPoint2_Max>0</BreakPoint2_Max>";
                BreakPoint2_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
                BreakPoint2_MaxEdit = "    <BreakPoint2_MaxEdit>0</BreakPoint2_MaxEdit>";
                BreakPoint2_ReferenceChannel = "    <BreakPoint2_ReferenceChannel />";
                BreakPoint2_Count = "    <BreakPoint2_Count>0</BreakPoint2_Count>";
                Exportable = "    <Exportable>true</Exportable>";
                Validated = "    <Validated>true</Validated>";
                BreakPoint1_FactoryName = "    <BreakPoint1_FactoryName>[0]</BreakPoint1_FactoryName>";
                BreakPoint2_FactoryName = "    <BreakPoint2_FactoryName>[0]</BreakPoint2_FactoryName>";
                IsArray = "    <IsArray>true</IsArray>";
                UseMaxSize = "    <UseMaxSize>true</UseMaxSize>";
                Open = "    <Open>false</Open>";
                Notes = "    <Notes />";
                BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
                BreakPoint2_Monotonicity = "    <BreakPoint2_Monotonicity>2</BreakPoint1_Monotonicity>";
                ELFVarType = "    <ELFVarType>2</ELFVarType>";
                Var_IncrementID = "    <Var_IncrementID />";
                Var_FormulaID = "    <Var_FormulaID />";
                BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";
                BreakPoint2_FormulaID = "    <BreakPoint2_FormulaID />";

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "#######0";
            string type_string = "65";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);
            if (P.parDecNum != 0)
            {
                format_string += ".";
                for (i = 0; i < P.parDecNum; i++)
                {
                    format_string += "0";
                }
            }
            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2 * P.parDim1);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "65";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2 * P.parDim1);
                    byte_dim_single_value = "1";
                    signed_string = "true";
                    type_string = "65";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2 * P.parDim1);
                    byte_dim_single_value = "2";
                    signed_string = "false";
                    type_string = "65";
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2 * P.parDim1);
                    byte_dim_single_value = "2";
                    signed_string = "true";
                    type_string = "65";
                    break;
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2 * P.parDim1);
                    byte_dim_single_value = "4";
                    signed_string = "false";
                    type_string = "65";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2 * P.parDim1);
                    byte_dim_single_value = "4";
                    signed_string = "true";
                    type_string = "65";
                    break;
                default:
                    byte_dim = Convert.ToString(P.parDim2 * P.parDim1);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "65";
                    break;
            }

            _name = P.parName;
            //******************************************************

            FactoryName = "    <FactoryName>" + P.parName + "[0]</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>" + P.parSource + "</GroupID>";
            CVType = "    <Type>"+ type_string +"</Type>";
            NByte = "    <NByte>" + byte_dim + "</NByte>";
            NByteSingleValue = "    <NByteSingleValue>" + byte_dim_single_value + "</NByteSingleValue>";
            Equispaced = "    <EquiSpaced>false</EquiSpaced>";
            Signed = "    <Signed>" + signed_string + "</Signed>";
            VarNameL0 = "    <VarNameL0>" + P.parAlias + "[0]</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.parName + "[0]</VarNameL1>";
            DescL0 = "    <DescL0>" + P.parDescription + "</DescL0>";
            DescL1 = "    <DescL1>" + P.parDescription + "</DescL1>";
            Var_Label = "    <Var_Label />";
            Var_ScalingID = "    <Var_ScalingID>" + scaling_string + "</Var_ScalingID>";
            Var_Format = "    <Var_Format>" + format_string + "</Var_Format>";
            Var_Unit = "    <Var_Unit>" + P.parUnit + "</Var_Unit>";
            Var_Min = "    <Var_Min>" + P.parMin + "</Var_Min>";
            Var_Max = "    <Var_Max>" + P.parMax + "</Var_Max>";
            Var_MinEdit = "    <Var_MinEdit>" + P.parMin + "</Var_MinEdit>";
            Var_MaxEdit = "    <Var_MaxEdit>" + P.parMax + "</Var_MaxEdit>";
            Var_ReferenceChannel = "    <Var_ReferenceChannel />";
            BreakPoint1_Label = "    <BreakPoint1_Label />";
            BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID />";
            BreakPoint1_Format = "    <BreakPoint1_Format />";
            BreakPoint1_Unit = "    <BreakPoint1_Unit />";
            BreakPoint1_Min = "    <BreakPoint1_Min>0</BreakPoint1_Min>";
            BreakPoint1_Max = "    <BreakPoint1_Max>0</BreakPoint1_Max>";
            BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
            BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>";
            BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel />";
            BreakPoint1_Count = "    <BreakPoint1_Count>0</BreakPoint1_Count>";
            BreakPoint2_Label = "    <BreakPoint2_Label />";
            BreakPoint2_ScalingID = "    <BreakPoint2_ScalingID />";
            BreakPoint2_Format = "    <BreakPoint2_Format />";
            BreakPoint2_Unit = "    <BreakPoint2_Unit />";
            BreakPoint2_Min = "    <BreakPoint2_Min>0</BreakPoint2_Min>";
            BreakPoint2_Max = "    <BreakPoint2_Max>0</BreakPoint2_Max>";
            BreakPoint2_MinEdit = "    <BreakPoint2_MinEdit>0</BreakPoint2_MinEdit>";
            BreakPoint2_MaxEdit = "    <BreakPoint2_MaxEdit>0</BreakPoint2_MaxEdit>";
            BreakPoint2_ReferenceChannel = "    <BreakPoint2_ReferenceChannel />";
            BreakPoint2_Count = "    <BreakPoint2_Count>0</BreakPoint2_Count>";
            Exportable = "    <Exportable>true</Exportable>";
            Validated = "    <Validated>true</Validated>";
            BreakPoint1_FactoryName = "    <BreakPoint1_FactoryName>" + P.parBreakpoint1 + "[0]</BreakPoint1_FactoryName>";
            BreakPoint2_FactoryName = "    <BreakPoint2_FactoryName>" + P.parBreakpoint2 + "[0]</BreakPoint2_FactoryName>";
            IsArray = "    <IsArray>true</IsArray>";
            UseMaxSize = "    <UseMaxSize>true</UseMaxSize>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
            BreakPoint2_Monotonicity = "    <BreakPoint2_Monotonicity>1</BreakPoint2_Monotonicity>";
            ELFVarType = "    <ELFVarType>2</ELFVarType>";
            Var_IncrementID = "    <Var_IncrementID />";
            Var_FormulaID = "    <Var_FormulaID />";
            BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";
            BreakPoint2_FormulaID = "    <BreakPoint2_FormulaID />";

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
                    Address = "    <Address>0x" + S.address + "</Address>";
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the Symbol database");
                }
                //                MessageBox.Show(_name + S.name + S.address);
            }

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <CalibrationMaps>");
            fileXML.WriteLine(FactoryName);
            fileXML.WriteLine(CalibrationTypeID);
            fileXML.WriteLine(Address);
            fileXML.WriteLine(OffsetByte);
            fileXML.WriteLine(OffsetBit);
            fileXML.WriteLine(GroupID);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(NByte);
            fileXML.WriteLine(NByteSingleValue);

            fileXML.WriteLine(Equispaced);

            fileXML.WriteLine(Signed);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);

            fileXML.WriteLine(Var_Label);
            fileXML.WriteLine(Var_ScalingID);
            fileXML.WriteLine(Var_Format);
            fileXML.WriteLine(Var_Unit);
            fileXML.WriteLine(Var_Min);
            fileXML.WriteLine(Var_Max);
            fileXML.WriteLine(Var_MinEdit);
            fileXML.WriteLine(Var_MaxEdit);
            fileXML.WriteLine(Var_ReferenceChannel);

            fileXML.WriteLine(BreakPoint1_Label);
            fileXML.WriteLine(BreakPoint1_ScalingID);
            fileXML.WriteLine(BreakPoint1_Format);
            fileXML.WriteLine(BreakPoint1_Unit);
            fileXML.WriteLine(BreakPoint1_Min);
            fileXML.WriteLine(BreakPoint1_Max);
            fileXML.WriteLine(BreakPoint1_MinEdit);
            fileXML.WriteLine(BreakPoint1_MaxEdit);
            fileXML.WriteLine(BreakPoint1_ReferenceChannel);
            fileXML.WriteLine(BreakPoint1_Count);

            fileXML.WriteLine(BreakPoint2_Label);
            fileXML.WriteLine(BreakPoint2_ScalingID);
            fileXML.WriteLine(BreakPoint2_Format);
            fileXML.WriteLine(BreakPoint2_Unit);
            fileXML.WriteLine(BreakPoint2_Min);
            fileXML.WriteLine(BreakPoint2_Max);
            fileXML.WriteLine(BreakPoint2_MinEdit);
            fileXML.WriteLine(BreakPoint2_MaxEdit);
            fileXML.WriteLine(BreakPoint2_ReferenceChannel);
            fileXML.WriteLine(BreakPoint2_Count);


            fileXML.WriteLine(Exportable);
            fileXML.WriteLine(Validated);
            fileXML.WriteLine(BreakPoint1_FactoryName);
            fileXML.WriteLine(BreakPoint2_FactoryName);
            fileXML.WriteLine(IsArray);
            fileXML.WriteLine(UseMaxSize);
            fileXML.WriteLine(Open);
            fileXML.WriteLine(Notes);
            fileXML.WriteLine(BreakPoint1_Monotonicity);
            fileXML.WriteLine(BreakPoint2_Monotonicity);
            fileXML.WriteLine(ELFVarType);
            fileXML.WriteLine(Var_IncrementID);
            fileXML.WriteLine(Var_FormulaID);
            fileXML.WriteLine(BreakPoint1_FormulaID);
            fileXML.WriteLine(BreakPoint2_FormulaID);
            fileXML.WriteLine("  </CalibrationMaps>");

            return (0);
        }

        public void Show()
        {
            MessageBox.Show("Maps");
        }
    }
    public class CalibrationSharedAxis
    {
        /*
        <CalibrationSharedAxes>
        <FactoryName>IgnCtl_tTCatxSACorr_Bk[0]</FactoryName>
        <CalibrationTypeID>LIE00</CalibrationTypeID>
        <Address>0x00A044BC</Address>
        <OffsetByte>17596</OffsetByte>
        <OffsetBit>0</OffsetBit>
        <GroupID>IGNCTL</GroupID>
        <Type>61</Type>
        <NByte>18</NByte>
        <NByteSingleValue>2</NByteSingleValue>
        <Signed>true</Signed>
        <Exportable>true</Exportable>
        <VarNameL0>IgnCtl_tTCatxSACorr_Bk[0]</VarNameL0>
        <VarNameL1>IgnCtl_tTCatxSACorr_Bk[0]</VarNameL1>
        <DescL0>IgnCtl_tTCatxSACorr_Bk[0]</DescL0>
        <DescL1>IgnCtl_tTCatxSACorr_Bk[0]</DescL1>
          
        <BreakPoint1_Label>TCat</BreakPoint1_Label>
        <BreakPoint1_ScalingID>S_1_1_0_0_1</BreakPoint1_ScalingID>
        <BreakPoint1_Format>###0</BreakPoint1_Format>
        <BreakPoint1_Unit>C</BreakPoint1_Unit>
        <BreakPoint1_Min>-32768</BreakPoint1_Min>
        <BreakPoint1_Max>32767</BreakPoint1_Max>
        <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>
        <BreakPoint1_MaxEdit>1023</BreakPoint1_MaxEdit>
        <BreakPoint1_ReferenceChannel>TCat[0]</BreakPoint1_ReferenceChannel>
        <BreakPoint1_Count>9</BreakPoint1_Count>
         
        <IsArray>true</IsArray>
        <UseMaxSize>true</UseMaxSize>
        <Validated>true</Validated>
        <Open>false</Open>
        <Notes />
        <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>
        <ELFVarType>2</ELFVarType>
        <BreakPoint1_FormulaID />
        </CalibrationSharedAxes>
        */

        internal string FactoryName;
        internal string CalibrationTypeID;
        internal string Address;
        internal string OffsetByte;
        internal string OffsetBit;
        internal string GroupID;
        internal string CVType;
        internal string NByte;
        internal string NByteSingleValue;
        internal string Signed;
        internal string Exportable;
        internal string VarNameL0;
        internal string VarNameL1;
        internal string DescL0;
        internal string DescL1;

        internal string BreakPoint1_Label;
        internal string BreakPoint1_ScalingID;
        internal string BreakPoint1_Format;
        internal string BreakPoint1_Unit;
        internal string BreakPoint1_Min;
        internal string BreakPoint1_Max;
        internal string BreakPoint1_MinEdit;
        internal string BreakPoint1_MaxEdit;
        internal string BreakPoint1_ReferenceChannel;
        internal string BreakPoint1_Count;

        internal string IsArray;
        internal string UseMaxSize;
        internal string Validated;
        internal string Open;
        internal string Notes;
        internal string BreakPoint1_Monotonicity;
        internal string ELFVarType;
        internal string BreakPoint1_FormulaID;

        public CalibrationSharedAxis()
        {
            FactoryName = "    <FactoryName>ZZZZZ</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>ENGST</GroupID>";
            CVType = "    <Type>61</Type>";
            NByte = "    <NByte>2</NByte>";
            NByteSingleValue = "    <NByteSingleValue>2</NByteSingleValue>";
            Signed = "    <Signed>false</Signed>";
            Exportable = "< Exportable > true </ Exportable >";

            VarNameL0 = "    <VarNameL0>ZZZZZ</VarNameL0>";
            VarNameL1 = "    <VarNameL1>ZZZZZ</VarNameL1>";
            DescL0 = "    <DescL0>rpm threshold above which Engine is On</DescL0>";
            DescL1 = "    <DescL1>rpm threshold above which Engine is On</DescL1>";
            
            BreakPoint1_Label = "    <BreakPoint1_Label />";
            BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID />";
            BreakPoint1_Format = "    <BreakPoint1_Format />";
            BreakPoint1_Unit = "    <BreakPoint1_Unit />";
            BreakPoint1_Min = "    <BreakPoint1_Min>0</BreakPoint1_Min>";
            BreakPoint1_Max = "    <BreakPoint1_Max>0</BreakPoint1_Max>";
            BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>0</BreakPoint1_MinEdit>";
            BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>0</BreakPoint1_MaxEdit>";
            BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel />";
            BreakPoint1_Count = "    <BreakPoint1_Count>0</BreakPoint1_Count>";

            IsArray = "    <IsArray>false</IsArray>";
            UseMaxSize = "    <UseMaxSize>false</UseMaxSize>";
            Validated = "    <Validated>true</Validated>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
            ELFVarType = "    <ELFVarType>2</ELFVarType>";
            BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "#######0";
            string type_string = "61";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);
            if (P.parDecNum != 0)
            {
                format_string += ".";
                for (i = 0; i < P.parDecNum; i++)
                {
                    format_string += "0";
                }
            }
            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "61";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "true";
                    type_string = "61";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "false";
                    type_string = "61";
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "true";
                    type_string = "61";
                    break;
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "false";
                    type_string = "61";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "true";
                    type_string = "61";
                    break;
                default:
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "false";
                    type_string = "61";
                    break;
            }

            _name = P.parName;
            //******************************************************

            FactoryName = "    <FactoryName>" + P.parName + "[0]</FactoryName>";
            CalibrationTypeID = "    <CalibrationTypeID>LIE00</CalibrationTypeID>";
            Address = "    <Address>0x00000000</Address>";
            OffsetByte = "    <OffsetByte>0</OffsetByte>";
            OffsetBit = "    <OffsetBit>0</OffsetBit>";
            GroupID = "    <GroupID>" + P.parSource + "</GroupID>";
            CVType = "    <Type>"+type_string+"</Type>";
            NByte = "    <NByte>" + byte_dim + "</NByte>";
            NByteSingleValue = "    <NByteSingleValue>" + byte_dim_single_value + "</NByteSingleValue>";
            Signed = "    <Signed>" + signed_string + "</Signed>";
            Exportable = "    <Exportable>true</Exportable>";
            VarNameL0 = "    <VarNameL0>" + P.parAlias + "[0]</VarNameL0>";
            VarNameL1 = "    <VarNameL1>" + P.parName + "</VarNameL1>";
            DescL0 = "    <DescL0>" + P.parDescription + "</DescL0>";
            DescL1 = "    <DescL1>" + P.parDescription + "</DescL1>";
            BreakPoint1_Label = "    <BreakPoint1_Label>"+P.parInputQuantity+ "</BreakPoint1_Label>";
            BreakPoint1_ScalingID = "    <BreakPoint1_ScalingID>" + scaling_string + "</BreakPoint1_ScalingID>" ;
            BreakPoint1_Format = "    <BreakPoint1_Format>"+ format_string +"</BreakPoint1_Format>";
            BreakPoint1_Unit = "    <BreakPoint1_Unit>" + P.parUnit + "</BreakPoint1_Unit>";
            BreakPoint1_Min = "    <BreakPoint1_Min>" + P.parMin + "</BreakPoint1_Min>";
            BreakPoint1_Max = "    <BreakPoint1_Max>" + P.parMax + "</BreakPoint1_Max>";
            BreakPoint1_MinEdit = "    <BreakPoint1_MinEdit>" + P.parMin + "</BreakPoint1_MinEdit>";
            BreakPoint1_MaxEdit = "    <BreakPoint1_MaxEdit>" + P.parMax + "</BreakPoint1_MaxEdit>";
            BreakPoint1_ReferenceChannel = "    <BreakPoint1_ReferenceChannel>"+P.parInputQuantity+"</BreakPoint1_ReferenceChannel>";
            BreakPoint1_Count = "    <BreakPoint1_Count>"+ Convert.ToString(P.parDim2) +"</BreakPoint1_Count>";
            IsArray = "    <IsArray>true</IsArray>";
            UseMaxSize = "    <UseMaxSize>true</UseMaxSize>";
            Validated = "    <Validated>true</Validated>";
            Open = "    <Open>false</Open>";
            Notes = "    <Notes />";
            BreakPoint1_Monotonicity = "    <BreakPoint1_Monotonicity>1</BreakPoint1_Monotonicity>";
            ELFVarType = "    <ELFVarType>2</ELFVarType>";
            BreakPoint1_FormulaID = "    <BreakPoint1_FormulaID />";

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
                    Address = "    <Address>0x" + S.address + "</Address>";
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the Symbol table");
                }
                //                MessageBox.Show(_name + S.name + S.address);
            }

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter fileXML)
        {
            fileXML.WriteLine("  <CalibrationSharedAxes>");
            fileXML.WriteLine(FactoryName);
            fileXML.WriteLine(CalibrationTypeID);
            fileXML.WriteLine(Address);
            fileXML.WriteLine(OffsetByte);
            fileXML.WriteLine(OffsetBit);
            fileXML.WriteLine(GroupID);
            fileXML.WriteLine(CVType);
            fileXML.WriteLine(NByte);
            fileXML.WriteLine(NByteSingleValue);

            fileXML.WriteLine(Signed);
            fileXML.WriteLine(Exportable);
            fileXML.WriteLine(VarNameL0);
            fileXML.WriteLine(VarNameL1);
            fileXML.WriteLine(DescL0);
            fileXML.WriteLine(DescL1);

            fileXML.WriteLine(BreakPoint1_Label);
            fileXML.WriteLine(BreakPoint1_ScalingID);
            fileXML.WriteLine(BreakPoint1_Format);
            fileXML.WriteLine(BreakPoint1_Unit);
            fileXML.WriteLine(BreakPoint1_Min);
            fileXML.WriteLine(BreakPoint1_Max);
            fileXML.WriteLine(BreakPoint1_MinEdit);
            fileXML.WriteLine(BreakPoint1_MaxEdit);
            fileXML.WriteLine(BreakPoint1_ReferenceChannel);
            fileXML.WriteLine(BreakPoint1_Count);


            fileXML.WriteLine(IsArray);
            fileXML.WriteLine(UseMaxSize);
            fileXML.WriteLine(Validated);
            fileXML.WriteLine(Open);
            fileXML.WriteLine(Notes);
            fileXML.WriteLine(BreakPoint1_Monotonicity);
            fileXML.WriteLine(ELFVarType);
            fileXML.WriteLine(BreakPoint1_FormulaID);
            fileXML.WriteLine("  </CalibrationSharedAxes>");

            return (0);
        }


        public void Show()
        {
            MessageBox.Show("SharedAxes");
        }

    }
}