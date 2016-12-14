using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication3
{
    public sealed class A2LGroup
    {
        internal string ID;
        internal List<string> MeasurementList;
        internal List<string> CharacteristicList;
        public A2LGroup()
        {
            ID = "";
        }
        public int upload(ref XLSECTSignal P)
        {
            return (0);
        }
        public int upload(ref XLSECTParameter P)
        {
            return (0);
        }

    }
    public sealed class AxisDescr
    {
        internal string Attribute;
        internal string InputQuantity;
        internal string Conversion;
        internal string MaxAxisPoints;
        internal string LowerLimit;
        internal string UpperLimit;
        internal string AxisPtsRef;

        public AxisDescr()
        {
            Attribute = "";
            InputQuantity = "";
            Conversion = "";
            MaxAxisPoints = "";
            LowerLimit = "";
            UpperLimit = "";
            AxisPtsRef = "AXIS_PTS_REF ";
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the conversion method  */
            file.WriteLine(spacing + "/begin AXIS_DESCR");
            file.WriteLine(spacing2 + Attribute);
            file.WriteLine(spacing2 + InputQuantity);
            file.WriteLine(spacing2 + Conversion);
            file.WriteLine(spacing2 + MaxAxisPoints);
            file.WriteLine(spacing2 + LowerLimit);
            file.WriteLine(spacing2 + UpperLimit);
            file.WriteLine(spacing2 + AxisPtsRef);
            file.WriteLine(spacing + "/end AXIS_DESCR");

            return 0;
        }
        public int GetLimits()
        {
            return (0);
        }
    }

    public sealed class Conversion
    {

//        /begin COMPU_METHOD
//            ANIN_InitDone.CONV
//			"ANIN_InitDone"
//			RAT_FUNC "%5.3"
//			""
//			COEFFS
//			0
//			-1.000000
//			0.000000
//			0
//			0.000000
//			-1.000000
//		/end COMPU_METHOD

        internal string Cumpu_method_ID;
        internal string comment;
        internal string rat_func;
        internal string measure_unit;
        internal string coeffs;
        internal string c1;
        internal string c2;
        internal string c3;
        internal string c4;
        internal string c5;
        internal string c6;


        public Conversion()
        {
            Cumpu_method_ID = "";
            comment = "";
            rat_func = "RAT_FUNC %5.3";
            measure_unit = "";
            coeffs = "COEFFS";
            c1 = "0";
            c2 = "0";
            c3 = "0";
            c4 = "0";
            c5 = "0";
            c6 = "0";

        }
        public int upload(string name, string k1, string k2, string k3, string k4, string format, string unit)
        {
            Cumpu_method_ID = name + ".CONV";
            comment = "\"Compu method for " + name +" \"";
            rat_func = "RAT_FUNC \""+ format + "\"";
            measure_unit = "\"" + unit + "\"";
            coeffs = "COEFFS";
            c1 = "0";
            c2 = "-"+k1;
            c3 = k2;
            c4 = "0";
            c5 = k3;
            c6 = "-"+k4;
            return (0);
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the conversion method  */
            file.WriteLine(spacing + "/begin COMPU_METHOD");
            file.WriteLine(spacing2 + Cumpu_method_ID);
            file.WriteLine(spacing2 + comment);
            file.WriteLine(spacing2 + rat_func);
            file.WriteLine(spacing2 + measure_unit);
            file.WriteLine(spacing2 + coeffs);
            file.WriteLine(spacing2 + c1);
            file.WriteLine(spacing2 + c2);
            file.WriteLine(spacing2 + c3);
            file.WriteLine(spacing2 + c4);
            file.WriteLine(spacing2 + c5);
            file.WriteLine(spacing2 + c6);
            file.WriteLine(spacing + "/end COMPU_METHOD");

            return 0;
        }

    }
    public sealed class Measurement
    {
        Conversion CompuMethod;
        //        /begin MEASUREMENT
        //            CAN_PRI_TCatL2
        //			"CAN_PRI_TCatL2"
        //			UWORD
        //           CAN_PRI_TCatL2.CONV
        //			0
        //			0
        //			0.000000
        //			65535.000000
        //			ECU_ADDRESS 0x40005adc
        //
        //            FORMAT "%6.0"
        //
        //            DISPLAY_IDENTIFIER CAN_PRI_TCatL2
        //
        //            READ_WRITE
        //			/begin IF_DATA ASAP1B_CCP
        //                KP_BLOB 0x0 0x40005adc 2
        //			/end IF_DATA
        //			/begin IF_DATA CANAPE_EXT
        //				100
        //              LINK_MAP CAN_PRI_TCatL2
        //				0x40005adc
        //				0x0
        //				0
        //				0x0
        //				1
        //				0x8f
        //				0x0
        //              DISPLAY 0 0.000000 7.999878
        //			/end IF_DATA
        //		/end MEASUREMENT

        internal string name;
        internal string long_identifier;
        internal string datatype;
        internal string conversion;
        internal string resolution;
        internal string accuracy;
        internal string lower_limit;
        internal string upper_limit;
        internal string ecu_address;
        internal string format;
        internal string display_identifier;
        internal string read_write;

//        /* for MEASUREMENT */
//       "KP_BLOB" struct {
//            uint;   /* Address extension of the online data
//            (only Low Byte significant) */
//            ulong;  /* Base address of the online data   */
//            ulong;  /* Number of Bytes belonging to the online data (1,2 or 4) */
//            taggedstruct {
//            ("RASTER" uchar )*;
//            /* Array of event channel initialization values */
//        };

        internal string asap1b_ccp;

        //      /*********************************************************/
        //		/*   CANape linker map references                        */
        //		/*********************************************************/
        //		"CANAPE_EXT" struct
        //       {
        //
        //          int;             /* version number */
        //			taggedstruct {
        //				"LINK_MAP" struct {
        //                  char[256];   /* segment name */
        //					long;        /* base address of the segment */
        //					uint;        /* address extension of the segment */
        //					uint;        /* flag: address is relative to DS */
        //					long;        /* offset of the segment address */
        //					uint;        /* datatypValid */
        //					uint;        /* enum datatyp */
        //					uint;        /* bit offset of the segment */
        //				};
        //				"DISPLAY" struct {
        //                  long;        /* display color */
        //					double;      /* minimal display value (phys)*/
        //					double;      /* maximal display value (phys)*/
        //				};
        //				"VIRTUAL_CONVERSION" struct {
        //                  char[256];   /* name of the conversion formula */
        //				};
        //			};
        //		};

        internal string version_number;
        internal string link_map;
        internal string base_address;
        internal string address_extension;
        internal string address_flag;
        internal string segment_offset;
        internal string datatyp_valid;
        internal string enum_datatyp;
        internal string segment_bit_offset;
        internal string display;

        public Measurement()
        {
            CompuMethod = new Conversion();
            name = "";
            long_identifier = "";
            datatype = "UWORD";
            conversion = "S_1_1_0_0_1";
            resolution = "0";
            accuracy = "0";
            lower_limit = "0.000000";
            upper_limit = "65535.000000";
            ecu_address = "ECU_ADDRESS 0x00000000";
            format = "FORMAT \" % 9.9\"";
            display_identifier = "DISPLAY_IDENTIFIER CAN_PRI_TCatL2";
            read_write = "READ_WRITE";
            asap1b_ccp = "KP_BLOB 0x0 0x00000000 2";
            version_number = "100";
            link_map = "LINK_MAP ";
            base_address = "0x00000000";
            address_extension = "0x0";
            address_flag = "0";
            segment_offset = "0x0";
            datatyp_valid = "1";
            enum_datatyp = "0x8f";
            segment_bit_offset = "0x0";
            display = "DISPLAY 0 0.000000 0.000000";
        }
        public int upload(ref XLSECTSignal P, bool isAnArray, int ind, ref Cont Cnr)
        {
            int i;
            string byte_dim = "1";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "%9";
            string type_string = "0";
            string _name = null;
            string i_hex;

            scaling_string = "S_1_" + P.sigKa + "_" + P.sigKb + "_" + P.sigKc + "_" + P.sigKd;

            if (P.sigDecNum != 0)
            {
                format_string += "."+P.sigDecNum;
            }

            switch (P.sigType)
            {
                case "UBYTE":
                    byte_dim = "1";
                    type_string = "0x87";
                    break;
                case "SBYTE":
                    byte_dim = "1";
                    type_string = "0xc7";
                    break;
                case "UWORD":
                    byte_dim = "2";
                    type_string = "0x8f";
                    break;
                case "SWORD":
                    byte_dim = "2";
                    type_string = "0xcf";
                    break;
                case "ULONG":
                    byte_dim = "4";
                    type_string = "0x9f";
                    break;
                case "SLONG":
                    byte_dim = "4";
                    type_string = "0x97";
                    break;
                default:
                    byte_dim = "1";
                    break;
            }

            _name = P.sigName;
            name = _name;

            long_identifier = "\"" + P.sigDescription + "\"";
            datatype = P.sigType;
            conversion = scaling_string;
            resolution = "0";
            accuracy = "0";
            lower_limit = P.sigMin;
            upper_limit = P.sigMax;
            ecu_address = "ECU_ADDRESS " + " 0x00000000";
            format = "FORMAT \"" + format_string +"\"";
            display_identifier = "DISPLAY_IDENTIFIER " + P.sigAlias;
            read_write = "READ_WRITE";
            asap1b_ccp = "KP_BLOB 0x0 0x00000000 " + byte_dim;
            version_number = "100";
            link_map = "LINK_MAP " +_name;
            base_address = "0x00000000";
            address_extension = "0x0";
            address_flag = "0";
            segment_offset = "0x0";
            datatyp_valid = "1";
            enum_datatyp = type_string;
            segment_bit_offset = "0x0";
            display = "DISPLAY 0 " + P.sigMin + " " + P.sigMax;


            if (isAnArray)
            {
                name = P.sigName + "[" + Convert.ToString(ind) + "]";
                long_identifier = "\"" + P.sigDescription +", element: " + Convert.ToString(ind) + "\"";
                display_identifier = "DISPLAY_IDENTIFIER " + P.sigAlias + "[" + Convert.ToString(ind) + "]";
            }

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolBssList.Find(x => x.name == _name);
                    if (isAnArray)
                    {
                        i = Convert.ToInt32(S.address, 16) + ind * Convert.ToInt32(byte_dim);
                        i_hex = i.ToString("X");
                        ecu_address = "ECU_ADDRESS 0x" + i_hex;
                        base_address = "0x" + i_hex;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + i_hex + " " + byte_dim;


                    }
                    else
                    {
                        ecu_address = "ECU_ADDRESS 0x" + S.address;
                        base_address = "0x" +S.address;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + " " + byte_dim;
                    }
                }
                catch
                {
                    MessageBox.Show(_name + "does not exist");
                }

            }
            if (0 == CompuMethod.upload(name, P.sigKa, P.sigKb, P.sigKc, P.sigKd, format_string, P.sigUnit))
            {
                conversion = CompuMethod.Cumpu_method_ID;
            }
            return (0);
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing =  "          ";
            string spacing2 = "                 ";
            string spacing3 = "                      ";

            /* Write the conversion method first */
            file.WriteLine(spacing + "/begin COMPU_METHOD");
            file.WriteLine(spacing2 + CompuMethod.Cumpu_method_ID);
            file.WriteLine(spacing2 + CompuMethod.comment);
            file.WriteLine(spacing2 + CompuMethod.rat_func);
            file.WriteLine(spacing2 + CompuMethod.measure_unit);
            file.WriteLine(spacing2 + CompuMethod.coeffs);
            file.WriteLine(spacing2 + CompuMethod.c1);
            file.WriteLine(spacing2 + CompuMethod.c2);
            file.WriteLine(spacing2 + CompuMethod.c3);
            file.WriteLine(spacing2 + CompuMethod.c4);
            file.WriteLine(spacing2 + CompuMethod.c5);
            file.WriteLine(spacing2 + CompuMethod.c6);
            file.WriteLine(spacing + "/end COMPU_METHOD");

            /* Write the measurement then */
            file.WriteLine(spacing + "/begin MEASUREMENT");
            file.WriteLine(spacing2 + name);
            file.WriteLine(spacing2 + long_identifier);
            file.WriteLine(spacing2 + datatype);
            file.WriteLine(spacing2 + conversion);
            file.WriteLine(spacing2 + resolution);
            file.WriteLine(spacing2 + accuracy);
            file.WriteLine(spacing2 + lower_limit);
            file.WriteLine(spacing2 + upper_limit);
            file.WriteLine(spacing2 + ecu_address);
            file.WriteLine(spacing2 + format);
            file.WriteLine(spacing2 + display_identifier);
            file.WriteLine(spacing2 + read_write);

            file.WriteLine(spacing2 + "/begin IF_DATA ASAP1B_CCP");
            file.WriteLine(spacing3 + asap1b_ccp);
            file.WriteLine(spacing2 + "/end IF_DATA");

            file.WriteLine(spacing2 + "/begin IF_DATA CANAPE_EXT");            
            file.WriteLine(spacing3 + version_number);
            file.WriteLine(spacing3 + link_map);
            file.WriteLine(spacing3 + base_address);
            file.WriteLine(spacing3 + address_extension);
            file.WriteLine(spacing3 + address_flag);
            file.WriteLine(spacing3 + segment_offset);
            file.WriteLine(spacing3 + datatyp_valid);
            file.WriteLine(spacing3 + enum_datatyp);
            file.WriteLine(spacing3 + segment_bit_offset);
            file.WriteLine(spacing3 + display);
            file.WriteLine(spacing2 + "/end IF_DATA");

            file.WriteLine(spacing + "/end MEASUREMENT");

            return 0;
        }
    }

    public sealed class ExtendedLimits
    {
        internal string LowerLimit;
        internal string UpperLimit;

        public ExtendedLimits ()
        {
            LowerLimit = "";
            UpperLimit = "";
        }
    }


    public sealed class CharacteristicGeneric
    {
        internal string name;
        internal string long_identifier;
        internal string characteristic_type;
        internal string address;
        internal string deposit;
        internal string max_diff;
        internal string conversion;
        internal string lower_limit;
        internal string upper_limit;
        internal string extended_limits;
        internal string ext_lower_limit;
        internal string ext_upper_limit;
        internal string format;
        internal string display_identifier;

        public CharacteristicGeneric()
        {
            name = "";
            long_identifier = "";
            characteristic_type = "VALUE";
            address = "0x00000000";
            deposit = "";
            max_diff = "0";
            conversion = "";
            lower_limit = "";
            upper_limit = "";
            extended_limits = "EXTENDED_LIMITS";
            ext_lower_limit = "";
            ext_upper_limit = "";
            format = "";
            display_identifier = "";
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the CHARACTERISTIC contents*/
            file.WriteLine(spacing2 + name);
            file.WriteLine(spacing2 + long_identifier);
            file.WriteLine(spacing2 + characteristic_type);
            file.WriteLine(spacing2 + address);
            file.WriteLine(spacing2 + deposit);
            file.WriteLine(spacing2 + max_diff);
            file.WriteLine(spacing2 + conversion);
            file.WriteLine(spacing2 + lower_limit);
            file.WriteLine(spacing2 + upper_limit);
            file.WriteLine(spacing2 + extended_limits);
            file.WriteLine(spacing2 + ext_lower_limit);
            file.WriteLine(spacing2 + ext_upper_limit);
            file.WriteLine(spacing2 + format);
            file.WriteLine(spacing2 + display_identifier);

            return (0);

        }
    }

    public sealed class ASAP1B_if_data
    {
        /* ASAP1B IF_DATA stuff */
        internal string asap1b_ccp;

        public ASAP1B_if_data()
        {
            asap1b_ccp = "";
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            file.WriteLine(spacing2 + "/begin IF_DATA ASAP1B_CCP");
            file.WriteLine(spacing3 + asap1b_ccp);
            file.WriteLine(spacing2 + "/end IF_DATA");

            return 0;
        }

    }
    public sealed class CanapeExt
    {
        /* CANAPE_EXT IF_DATA stuff */
        internal string version_number;
        internal string link_map;
        internal string base_address;
        internal string address_extension;
        internal string address_flag;
        internal string segment_offset;
        internal string datatyp_valid;
        internal string enum_datatyp;
        internal string segment_bit_offset;
        internal string display;

        public CanapeExt()
        {
            /* CANAPE_EXT IF_DATA stuff */
            version_number = "100";
            link_map = "";
            base_address = "";
            address_extension = "";
            address_flag = "";
            segment_offset = "";
            datatyp_valid = "";
            enum_datatyp = "";
            segment_bit_offset = "";
            display = "";
        }

        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            file.WriteLine(spacing2 + "/begin IF_DATA CANAPE_EXT");
            file.WriteLine(spacing3 + version_number);
            file.WriteLine(spacing3 + link_map);
            file.WriteLine(spacing3 + base_address);
            file.WriteLine(spacing3 + address_extension);
            file.WriteLine(spacing3 + address_flag);
            file.WriteLine(spacing3 + segment_offset);
            file.WriteLine(spacing3 + datatyp_valid);
            file.WriteLine(spacing3 + enum_datatyp);
            file.WriteLine(spacing3 + segment_bit_offset);
            file.WriteLine(spacing3 + display);
            file.WriteLine(spacing2 + "/end IF_DATA");

            return 0;
        }

    }
    public sealed class CharacteristicValue
    {
        Conversion CompuMethod;

        //        /begin CHARACTERISTIC
        //            KDgTps1FctValTh
        //			"KDgTps1FctValTh"
        //			VALUE
        //			0xa00820
        //			ValueUnsignedWord
        //			0
        //			KDgTps1FctValTh.val.CONV
        //			0.000000
        //			65535.000000
        //			EXTENDED_LIMITS
        //			0.000000
        //			65535.000000
        //			FORMAT "%5.3"
        //			DISPLAY_IDENTIFIER KDgTps1FctValTh
        //			/begin IF_DATA ASAP1B_CCP
        //               DP_BLOB 0x0 0xa00820 2
        //			/end IF_DATA
        //			/begin IF_DATA CANAPE_EXT
        //				100
        //				LINK_MAP KDgTps1FctValTh
        //				0xa00820
        //				0x0
        //				0
        //				0x0
        //				1
        //				0x8f
        //				0x0
        //              DISPLAY 0 0.000000 65535.000000
        //			/end IF_DATA
        //		/end CHARACTERISTIC
        internal string name;
        internal string long_identifier;
        internal string characteristic_type;
        internal string address;
        internal string deposit;
        internal string max_diff;
        internal string conversion;
        internal string lower_limit;
        internal string upper_limit;
        internal string extended_limits;
        internal string ext_lower_limit;
        internal string ext_upper_limit;
        internal string format;
        internal string display_identifier;

        /* ASAP1B IF_DATA stuff */
        internal string asap1b_ccp;

        /* CANAPE_EXT IF_DATA stuff */
        internal string version_number;
        internal string link_map;
        internal string base_address;
        internal string address_extension;
        internal string address_flag;
        internal string segment_offset;
        internal string datatyp_valid;
        internal string enum_datatyp;
        internal string segment_bit_offset;
        internal string display;


        public CharacteristicValue()
        {
            CompuMethod = new Conversion();

            name = "";
			long_identifier = "";
			characteristic_type = "VALUE";
			address = "0x00000000";
            deposit = "";
            max_diff = "0";
            conversion = "";
			lower_limit = "";
			upper_limit = "";
            extended_limits = "EXTENDED_LIMITS";
            ext_lower_limit = "";
            ext_upper_limit = "";
            format = "";
			display_identifier = "";
			
			/* ASAP1B IF_DATA stuff */
			asap1b_ccp = "";
			
			/* CANAPE_EXT IF_DATA stuff */
			version_number = "";
			link_map = "";
			base_address = "";
			address_extension = "";
			address_flag = "";
			segment_offset = "";
			datatyp_valid = "";
			enum_datatyp = "";
			segment_bit_offset = "";
			display = "";

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "%9";
            string type_string = "61";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);

            if (P.parDecNum != 0)
            {
                format_string += "." + P.parDecNum;
            }

            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "ValueUnsignedByte";
                    type_string = "0x87";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "ValueSignedByte";
                    type_string = "0xc7";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "ValueUnsignedWord";
                    type_string = "0x8f";
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "ValueSignedWord";
                    type_string = "0xcf";
                    break;
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueUnsignedLong";
                    type_string = "0x9f";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueSignedLong";
                    type_string = "0xdf";
                    break;
                default:
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "ValueUnsignedByte";
                    type_string = "0x87";
                    break;
            }

            _name = P.parName;
            //******************************************************

			name = _name;
			long_identifier = "\"" + P.parDescription + "\"";
			characteristic_type = "VALUE";
			address = "0x00000000";
            deposit = signed_string;
            max_diff = "0";
            conversion = scaling_string;
			lower_limit = P.parMin;
			upper_limit = P.parMax;
            extended_limits = "EXTENDED_LIMITS";
            ext_lower_limit = P.parMin;
            ext_upper_limit = P.parMax;
            format = "FORMAT \"" + format_string + "\"";
            display_identifier = "DISPLAY_IDENTIFIER " + P.parAlias;

            /* ASAP1B IF_DATA stuff */
            asap1b_ccp = "KP_BLOB 0x0 0x00000000 " + byte_dim;

            /* CANAPE_EXT IF_DATA stuff */
            version_number = "100";
            link_map = "LINK_MAP " + _name;
            base_address = "0x00000000";
            address_extension = "0x0";
            address_flag = "0";
            segment_offset = "0x0";
            datatyp_valid = "1";
            enum_datatyp = type_string;
            segment_bit_offset = "0x0";
            display = "DISPLAY 0 " + P.parMin + " " + P.parMax;

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
#if CALIBRATION_ARRAYS
                    if (isAnArray)
                    {
                        i = Convert.ToInt32(S.address, 16) + ind * Convert.ToInt32(byte_dim);
                        i_hex = i.ToString("X");
                        ecu_address = "ECU_ADDRESS 0x" + i_hex;
                        base_address = "0x" + i_hex;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + i_hex + " " + byte_dim;

                    }
                    else
                    {
                        ecu_address = "ECU_ADDRESS 0x" + S.address;
                        base_address = "0x" + S.address;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + byte_dim;
                    }
#else
                    address = "0x" + S.address;
                    base_address = "0x" + S.address;
                    asap1b_ccp = "DP_BLOB 0x0 0x" + S.address + " " +byte_dim;

#endif
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the symbol table");
                }

            }

            if (0 == CompuMethod.upload(name, Convert.ToString(P.parKa), Convert.ToString(P.parKb), Convert.ToString(P.parKc), Convert.ToString(P.parKd), format_string, P.parUnit))
            {
                conversion = CompuMethod.Cumpu_method_ID;
            }

            return (0);
        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing =  "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the conversion method first */
            file.WriteLine(spacing + "/begin COMPU_METHOD");
            file.WriteLine(spacing2 + CompuMethod.Cumpu_method_ID);
            file.WriteLine(spacing2 + CompuMethod.comment);
            file.WriteLine(spacing2 + CompuMethod.rat_func);
            file.WriteLine(spacing2 + CompuMethod.measure_unit);
            file.WriteLine(spacing2 + CompuMethod.coeffs);
            file.WriteLine(spacing2 + CompuMethod.c1);
            file.WriteLine(spacing2 + CompuMethod.c2);
            file.WriteLine(spacing2 + CompuMethod.c3);
            file.WriteLine(spacing2 + CompuMethod.c4);
            file.WriteLine(spacing2 + CompuMethod.c5);
            file.WriteLine(spacing2 + CompuMethod.c6);
            file.WriteLine(spacing + "/end COMPU_METHOD");
            /* Write the measurement then */
            file.WriteLine(spacing + "/begin CHARACTERISTIC");
            file.WriteLine(spacing2 + name);
            file.WriteLine(spacing2 + long_identifier);
            file.WriteLine(spacing2 + characteristic_type);
            file.WriteLine(spacing2 + address);
            file.WriteLine(spacing2 + deposit);
            file.WriteLine(spacing2 + max_diff);
            file.WriteLine(spacing2 + conversion);
            file.WriteLine(spacing2 + lower_limit);
            file.WriteLine(spacing2 + upper_limit);
            file.WriteLine(spacing2 + extended_limits);
            file.WriteLine(spacing2 + ext_lower_limit);
            file.WriteLine(spacing2 + ext_upper_limit);
            file.WriteLine(spacing2 + format);
            file.WriteLine(spacing2 + display_identifier);

            file.WriteLine(spacing2 + "/begin IF_DATA ASAP1B_CCP");
            file.WriteLine(spacing3 + asap1b_ccp);
            file.WriteLine(spacing2 + "/end IF_DATA");

            file.WriteLine(spacing2 + "/begin IF_DATA CANAPE_EXT");
            file.WriteLine(spacing3 + version_number);
            file.WriteLine(spacing3 + link_map);
            file.WriteLine(spacing3 + base_address);
            file.WriteLine(spacing3 + address_extension);
            file.WriteLine(spacing3 + address_flag);
            file.WriteLine(spacing3 + segment_offset);
            file.WriteLine(spacing3 + datatyp_valid);
            file.WriteLine(spacing3 + enum_datatyp);
            file.WriteLine(spacing3 + segment_bit_offset);
            file.WriteLine(spacing3 + display);
            file.WriteLine(spacing2 + "/end IF_DATA");

            file.WriteLine(spacing + "/end CHARACTERISTIC");

            return 0;
        }
    }

    public sealed class CharacteristicCurve
    {
        //        /begin CHARACTERISTIC
        //          Air_fPBaroCor_V[0]
        //			"1D Map to obtain the correction factor for the Relativa Air Charge based on PBaro"
        //			CURVE
        //			0xa00a80
        //			Curve_SignedWord_Com_Ax
        //			0
        //			Air_fPBaroCor_V[0].val.CONV
        //			-8.000000
        //			7.999756
        //			EXTENDED_LIMITS
        //			-8.000000
        //			7.999756
        //			FORMAT "%5.3"
        //			DISPLAY_IDENTIFIER Air_fPBaroCor_V[0]
        //			/begin AXIS_DESCR
        //              COM_AXIS
        //                NO_INPUT_QUANTITY
        //                NO_COMPU_METHOD
        //				8
        //				0
        //				4000
        //				AXIS_PTS_REF Air_fPBaroCor_Bk[0]
        //			/end AXIS_DESCR
        //			/begin IF_DATA ASAP1B_CCP
        //              DP_BLOB 0x0 0xa00a80 16
        //			/end IF_DATA
        //			/begin IF_DATA CANAPE_EXT
        //				100
        //				LINK_MAP Air_fPBaroCor_V._0_
        //				0xa00a80
        //				0x0
        //				0
        //				0x0
        //				1
        //				0xff
        //				0x0
        //				DISPLAY 0 -8.000000 7.999756
        //			/end IF_DATA
        //		/end CHARACTERISTIC

        Conversion CompuMethod;
        CharacteristicGeneric ChGen;
        ASAP1B_if_data AData;
        AxisDescr AD;
        CanapeExt CExt;

        public CharacteristicCurve()
        {
            CompuMethod= new Conversion();
            ChGen = new CharacteristicGeneric();
            AData = new ASAP1B_if_data();
            AD = new AxisDescr();
            CExt = new CanapeExt();

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "%9";
            string type_string = "61";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);

            if (P.parDecNum != 0)
            {
                format_string += "." + P.parDecNum;
            }

            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "Curve_Byte_Com_Ax";
                    type_string = "0xb7";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "Curve_SignedByte_Com_Ax";
                    type_string = "0xf7";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "Curve_Word_Com_Ax";
                    type_string = "0xbf";
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "Curve_SignedWord_Com_Ax";
                    type_string = "0xff";
                    break;
#if BOH
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueUnsignedLong";
                    type_string = "0x9f";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueSignedLong";
                    type_string = "0xdf";
                    break;
#endif
                default:
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "Curve_SignedWord_Com_Ax";
                    type_string = "0xff";
                    break;
            }

            _name = P.parName;
            //******************************************************

            ChGen.name = _name+ "[0]";
            ChGen.long_identifier = "\"" + P.parDescription + "\"";
            ChGen.characteristic_type = "CURVE";
            ChGen.address = "0x00000000";
            ChGen.deposit = signed_string;
            ChGen.max_diff = "0" ;
            ChGen.conversion = scaling_string;
            ChGen.lower_limit = P.parMin;
            ChGen.upper_limit = P.parMax;
            ChGen.extended_limits = "EXTENDED_LIMITS";
            ChGen.ext_lower_limit = P.parMin;
            ChGen.ext_upper_limit = P.parMax;
            ChGen.format = "FORMAT \"" + format_string + "\"";
            ChGen.display_identifier = "DISPLAY_IDENTIFIER " + P.parAlias;

            /* ASAP1B IF_DATA stuff */
            AData.asap1b_ccp = "DP_BLOB 0x0 0x00000000 " + byte_dim;

            /* CANAPE_EXT IF_DATA stuff */
            CExt.version_number = "100";
            CExt.link_map = "LINK_MAP " + _name + "._0_";
            CExt.base_address = "0x00000000";
            CExt.address_extension = "0x0";
            CExt.address_flag = "0";
            CExt.segment_offset = "0x0";
            CExt.datatyp_valid = "1";
            CExt.enum_datatyp = type_string;
            CExt.segment_bit_offset = "0x0";
            CExt.display = "DISPLAY 0 " + P.parMin + " " + P.parMax;

            /* Axis Descr */
            AD.Attribute = "COM_AXIS";
            AD.InputQuantity = "NO_INPUT_QUANTITY";
            AD.Conversion = "NO_COMPU_METHOD";
            AD.MaxAxisPoints = Convert.ToString(P.parDim2);
            AD.LowerLimit = "0";
            AD.UpperLimit = "0"; ////////////////////////////////////////////////////DA FARE, LISTA ASSI PER METTERE QUI I LIMITI
            AD.AxisPtsRef = "AXIS_PTS_REF " + P.parBreakpoint1 + "[0]";
            AD.GetLimits();

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
#if CALIBRATION_ARRAYS
                    if (isAnArray)
                    {
                        i = Convert.ToInt32(S.address, 16) + ind * Convert.ToInt32(byte_dim);
                        i_hex = i.ToString("X");
                        ecu_address = "ECU_ADDRESS 0x" + i_hex;
                        base_address = "0x" + i_hex;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + i_hex + " " + byte_dim;

                    }
                    else
                    {
                        ecu_address = "ECU_ADDRESS 0x" + S.address;
                        base_address = "0x" + S.address;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + byte_dim;
                    }
#else
                    ChGen.address = "0x" + S.address;
                    CExt.base_address = "0x" + S.address;
                    AData.asap1b_ccp = "DP_BLOB 0x0 0x" + S.address + " " + byte_dim;

#endif
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the symbol table");
                }

            }

            if (0 == CompuMethod.upload(ChGen.name, Convert.ToString(P.parKa), Convert.ToString(P.parKb), Convert.ToString(P.parKc), Convert.ToString(P.parKd), format_string, P.parUnit))
            {
                ChGen.conversion = CompuMethod.Cumpu_method_ID;
            }

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
                string spacing = "          ";
                string spacing2 = "                ";
                string spacing3 = "                      ";

                /* Write the conversion method first */

                CompuMethod.AppendToFile(ref file);

                file.WriteLine(spacing + "/begin CHARACTERISTIC");

                ChGen.AppendToFile(ref file);
                AD.AppendToFile(ref file);
                AData.AppendToFile(ref file);
                CExt.AppendToFile(ref file);

                file.WriteLine(spacing + "/end CHARACTERISTIC");

                return 0;
        }
    }
    public sealed class CharacteristicMap
    {
        Conversion CompuMethod;
        CharacteristicGeneric ChGen;
        ASAP1B_if_data AData;
        AxisDescr AD;
        AxisDescr AD2;
        CanapeExt CExt;

        public CharacteristicMap()
        {
            CompuMethod = new Conversion();
            ChGen = new CharacteristicGeneric();
            AData = new ASAP1B_if_data();
            AD = new AxisDescr();
            AD2 = new AxisDescr();
            CExt = new CanapeExt();

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            string byte_dim = "1";
            string byte_dim_single_value = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "%9";
            string type_string = "61";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);

            if (P.parDecNum != 0)
            {
                format_string += "." + P.parDecNum;
            }

            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "Map_Byte_Com_Ax";
                    type_string = "0x3a7"; /* Verified */
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "1";
                    signed_string = "Map_SignedByte_Com_Ax";
                    type_string = "0xf7";  /* Unverified */
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "Map_Word_Com_Ax";
                    type_string = "0x17bf"; /* Unverified */
                    break;
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "Map_SignedWord_Com_Ax";
                    type_string = "0x17ff"; /* Verified */
                    break;
#if BOH
                case "ULONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueUnsignedLong";
                    type_string = "0x9f";
                    break;
                case "SLONG":
                    byte_dim = Convert.ToString(4 * P.parDim2);
                    byte_dim_single_value = "4";
                    signed_string = "ValueSignedLong";
                    type_string = "0xdf";
                    break;
#endif
                default:
                    byte_dim = Convert.ToString(P.parDim2);
                    byte_dim_single_value = "2";
                    signed_string = "Map_SignedWord_Com_Ax";
                    type_string = "0x17ff"; /* Most common */
                    break;
            }

            _name = P.parName;
            //******************************************************

            ChGen.name = _name + "[0]";
            ChGen.long_identifier = "\"" + P.parDescription + "\"";
            ChGen.characteristic_type = "CURVE";
            ChGen.address = "0x00000000";
            ChGen.deposit = signed_string;
            ChGen.max_diff = "0";
            ChGen.conversion = scaling_string;
            ChGen.lower_limit = P.parMin;
            ChGen.upper_limit = P.parMax;
            ChGen.extended_limits = "EXTENDED_LIMITS";
            ChGen.ext_lower_limit = P.parMin;
            ChGen.ext_upper_limit = P.parMax;
            ChGen.format = "FORMAT \"" + format_string + "\"";
            ChGen.display_identifier = "DISPLAY_IDENTIFIER " + P.parAlias;

            /* ASAP1B IF_DATA stuff */
            AData.asap1b_ccp = "DP_BLOB 0x0 0x00000000 " + byte_dim;

            /* CANAPE_EXT IF_DATA stuff */
            CExt.version_number = "100";
            CExt.link_map = "LINK_MAP " + _name + "._0_";
            CExt.base_address = "0x00000000";
            CExt.address_extension = "0x0";
            CExt.address_flag = "0";
            CExt.segment_offset = "0x0";
            CExt.datatyp_valid = "1";
            CExt.enum_datatyp = type_string;
            CExt.segment_bit_offset = "0x0";
            CExt.display = "DISPLAY 0 " + P.parMin + " " + P.parMax;

            /* Axis 1 Descr */
            AD.Attribute = "COM_AXIS";
            AD.InputQuantity = "NO_INPUT_QUANTITY";
            AD.Conversion = "NO_COMPU_METHOD";
            AD.MaxAxisPoints = Convert.ToString(P.parDim1);
            AD.LowerLimit = "0";
            AD.UpperLimit = "0";
            AD.AxisPtsRef = "AXIS_PTS_REF " + P.parBreakpoint1 + "[0]";
            AD.GetLimits();

            /* Axis 2 Descr */
            AD2.Attribute = "COM_AXIS";
            AD2.InputQuantity = "NO_INPUT_QUANTITY";
            AD2.Conversion = "NO_COMPU_METHOD";
            AD2.MaxAxisPoints = Convert.ToString(P.parDim2);
            AD2.LowerLimit = "0";
            AD2.UpperLimit = "0"; 
            AD2.AxisPtsRef = "AXIS_PTS_REF " + P.parBreakpoint2 + "[0]";
            AD2.GetLimits();

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
#if CALIBRATION_ARRAYS
                    if (isAnArray)
                    {
                        i = Convert.ToInt32(S.address, 16) + ind * Convert.ToInt32(byte_dim);
                        i_hex = i.ToString("X");
                        ecu_address = "ECU_ADDRESS 0x" + i_hex;
                        base_address = "0x" + i_hex;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + i_hex + " " + byte_dim;

                    }
                    else
                    {
                        ecu_address = "ECU_ADDRESS 0x" + S.address;
                        base_address = "0x" + S.address;
                        asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + byte_dim;
                    }
#else
                    ChGen.address = "0x" + S.address;
                    CExt.base_address = "0x" + S.address;
                    AData.asap1b_ccp = "DP_BLOB 0x0 0x" + S.address + " " + byte_dim;

#endif
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the symbol table");
                }

            }

            if (0 == CompuMethod.upload(ChGen.name, Convert.ToString(P.parKa), Convert.ToString(P.parKb), Convert.ToString(P.parKc), Convert.ToString(P.parKd), format_string, P.parUnit))
            {
                ChGen.conversion = CompuMethod.Cumpu_method_ID;
            }

            return (0);

        }

        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the conversion method first */

            CompuMethod.AppendToFile(ref file);

            file.WriteLine(spacing + "/begin CHARACTERISTIC");

            ChGen.AppendToFile(ref file);
            AD.AppendToFile(ref file);
            AD2.AppendToFile(ref file);
            AData.AppendToFile(ref file);
            CExt.AppendToFile(ref file);

            file.WriteLine(spacing + "/end CHARACTERISTIC");

            return 0;
        }

    }
    public sealed class Axis_Pts
    {
        //        /begin COMPU_METHOD
        //          AirCtl_MapReqConf00_Bk1[0].bkpX.CONV "Conversione bkpX"
        //			RAT_FUNC "%5.0"
        //			"Rpm"
        //			COEFFS
        //			0
        //			-1.000000
        //			0.000000
        //			0
        //			0.000000
        //			-1.000000
        //		/end COMPU_METHOD
        //		/begin AXIS_PTS
        //          AirCtl_MapReqConf00_Bk1[0]                                  Name
        //			"AirCtl_MapReqConf00_Bk1[0]"                                longIdentifier
        //			0xa08928                                                    Address
        //			Rpm                                                         InputQuantity
        //          Com_Axis_SignedWord                                         Deposit
        //			0                                                           MaxDiff
        //          AirCtl_MapReqConf00_Bk1[0].bkpX.CONV                        Conversion
        //			24                                                          MaxAxisPoints
        //			0                                                           LowerLimit
        //			12000                                                       UpperLimit
        //			EXTENDED_LIMITS
        //			-32768
        //			32767
        //			FORMAT "%5.0"
        //			DEPOSIT ABSOLUTE
        //		/end AXIS_PTS

        Conversion CompuMethod;
        internal string name;
        internal string long_identifier;
        internal string address;
        internal string input_quantity;
        internal string deposit;
        internal string max_diff;
        internal string conversion;
        internal string max_axis_points;
        internal string lower_limit;
        internal string upper_limit;

        internal string extended_limits;
        internal string ext_lower_limit;
        internal string ext_upper_limit;

        internal string format;

        internal string DEPOSIT;

        public Axis_Pts()
        {
            CompuMethod = new Conversion();
            name = "";
            long_identifier = "";
            address = "0x00000000";
            input_quantity = "";
            deposit = "";
            max_diff = "0";
            conversion = "";
            max_axis_points = "";
            lower_limit = "";
            upper_limit = "";
            extended_limits = "EXTENDED_LIMITS";
            ext_lower_limit = "";
            ext_upper_limit = "";
            format = "";
            DEPOSIT = "DEPOSIT ABSOLUTE";

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            string byte_dim = "1";
            string signed_string = "false";
            string scaling_string = "S_1_1_0_0_1";
            string format_string = "%9";
            string _name = null;

            scaling_string = "S_1_" + Convert.ToString(P.parKa) + "_" + Convert.ToString(P.parKb) + "_" + Convert.ToString(P.parKc) + "_" + Convert.ToString(P.parKd);

            if (P.parDecNum != 0)
            {
                format_string += "." + P.parDecNum;
            }

            switch (P.parType)
            {
                case "UBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    signed_string = "Com_Axis_Byte";
                    break;
                case "SBYTE":
                    byte_dim = Convert.ToString(P.parDim2);
                    signed_string = "Com_Axis_SignedByte";
                    break;
                case "UWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    signed_string = "Com_Axis_Word";
                    break;
                default:
                case "SWORD":
                    byte_dim = Convert.ToString(2 * P.parDim2);
                    signed_string = "Com_Axis_SignedWord";
                    break;
            }

            _name = P.parName;
            //******************************************************

            name = _name + "[0]";
            long_identifier = "\"" + P.parDescription + "\"";
            address = "0x00000000";
            input_quantity = P.parInputQuantity;
            deposit = signed_string;
            max_diff = "0";
            conversion = scaling_string;
            max_axis_points = Convert.ToString(P.parDim2);
            lower_limit = P.parMin;
            upper_limit = P.parMax;
            extended_limits = "EXTENDED_LIMITS";
            ext_lower_limit = P.parMin;
            ext_upper_limit = P.parMax;
            format = "FORMAT \"" + format_string + "\"";

            if (Cnr.SymbolTableExists == true)
            {
                Symbol S = new Symbol();
                try
                {
                    S = Cnr.SymbolCalList.Find(x => x.name == _name);
                    address = "0x" + S.address;
                }
                catch
                {
                    MessageBox.Show(_name + " does not exist in the symbol table");
                }

            }

            if (0 == CompuMethod.upload(name, Convert.ToString(P.parKa), Convert.ToString(P.parKb), Convert.ToString(P.parKc), Convert.ToString(P.parKd), format_string, P.parUnit))
            {
                conversion = CompuMethod.Cumpu_method_ID;
            }

            return (0);

        }
        public int AppendToFile(ref System.IO.StreamWriter file)
        {
            string spacing = "          ";
            string spacing2 = "                ";
            string spacing3 = "                      ";

            /* Write the conversion method first */

            CompuMethod.AppendToFile(ref file);

            file.WriteLine(spacing + "/begin AXIS_PTS");
            file.WriteLine(spacing2 + name);
            file.WriteLine(spacing2 + long_identifier);
            file.WriteLine(spacing2 + address);
            file.WriteLine(spacing2 + input_quantity);
            file.WriteLine(spacing2 + deposit);
            file.WriteLine(spacing2 + max_diff);
            file.WriteLine(spacing2 + conversion);
            file.WriteLine(spacing2 + max_axis_points);
            file.WriteLine(spacing2 + lower_limit);
            file.WriteLine(spacing2 + upper_limit);
            file.WriteLine(spacing2 + extended_limits);
            file.WriteLine(spacing2 + ext_lower_limit);
            file.WriteLine(spacing2 + ext_upper_limit);
            file.WriteLine(spacing2 + format);
            file.WriteLine(spacing2 + DEPOSIT);
            file.WriteLine(spacing + "/end AXIS_PTS");

            return 0;
        }

    }
}
