using System;

namespace WindowsFormsApplication3
{

    public class XLSECTParameter
    {
        private string parName;
        private string parSource;
        private string parType;
        private int parDecNum;
        private string parUnit;
        private int parMin;
        private int parMax;
        private int parKa;
        private int parKb;
        private int parKc;
        private int parKd;
        private int parKk;
        private string parAlias;
        private string parDescription;
        private string parDimension;
        private string parBreakpoint1;
        private string parBreakpoint2;
        private string parInputQuantity;
        private string parStatesRow;

        public XLSECTParameter()
        {
            parName = "";
            parSource = "";
            parType = "";
            parDecNum = 0;
            parUnit = "";
            parMin = 0;
            parMax = 0;
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
        private string FactoryName;
        private string CalibrationTypeID;
        private string Address;
        private string OffsetByte;
        private string OffsetBit;
        private string GroupID;
        private string CVType;
        private string NByte;
        private string NByteSingleValue;
        private string NBit;
        private string Signed;
        private string VarNameL0;
        private string VarNameL1;
        private string DescL0;
        private string DescL1;
        private string Var_ScalingID;
        private string Var_Format;
        private string Var_Unit;
        private string Var_Min;
        private string Var_Max;
        private string Var_MinEdit;
        private string Var_MaxEdit;
        private string Exportable;
        private string Validated;
        private string SubType;
        private string IsArray;
        private string UseMaxSize;
        private string Open;
        private string Notes;
        private string ELFVarType;
        private string Var_FormulaID;

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
    }
}