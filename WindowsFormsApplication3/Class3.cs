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
            rat_func = "rat_func %5.3";
            measure_unit = "";
            coeffs = "coeffs";
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
                        asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + byte_dim;
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
        internal string datatype;
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
			datatype = "";
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
			datatype = signed_string;
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
                    address = "ECU_ADDRESS 0x" + S.address;
                    base_address = "0x" + S.address;
                    asap1b_ccp = "KP_BLOB 0x0 0x" + S.address + byte_dim;

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
            string spacing = "          ";
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
            file.WriteLine(spacing + "/begin CHARACTERISTIC");
            file.WriteLine(spacing2 + name);
            file.WriteLine(spacing2 + long_identifier);
            file.WriteLine(spacing2 + characteristic_type);
            file.WriteLine(spacing2 + address);
            file.WriteLine(spacing2 + datatype);
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
        public CharacteristicCurve()
        {

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            return (0);

        }
    }
    public sealed class CharacteristicMap
    {
        public CharacteristicMap()
        {

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            return (0);

        }
    }
    public sealed class Axis_Pts
    {
        public Axis_Pts()
        {

        }
        public int upload(ref XLSECTParameter P, ref Cont Cnr)
        {
            return (0);

        }
    }
}
