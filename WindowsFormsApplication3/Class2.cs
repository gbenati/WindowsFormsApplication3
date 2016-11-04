using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication3
{
    public class Cont
    {
        internal List<Symbol> SymbolBssList;
        internal List<Symbol> SymbolCalList;
        internal bool SymbolTableExists;

        public Cont()
        {
            SymbolBssList = new List<Symbol>();
            SymbolCalList = new List<Symbol>();
            SymbolTableExists = false;
        }

    }
    class RawSymbol
    {
        internal string field1;
        internal string field2;
        internal string field3;
        internal string field4;
        internal string field5;
        internal string field6;

        public RawSymbol(ref string line)
        {

            Char[] delimiter = { ' ', '\t' };

            if (line != null)
            {

                string[] subs = line.Split(delimiter, System.StringSplitOptions.RemoveEmptyEntries);

                if ((subs.Length != 0) && (line.Length != 0))
                {
                    if ((line[0] != ' ') && (subs[0] != "") && (subs[0] != null))
                    {
                        char c1 = subs[0][0];
                        if ((c1 < '0') || (c1 > '9') || (subs[0] == "00000000") || (subs.Length != 6))
                        {
                            field6 = "foo";
                        }
                        else
                        {
                            field1 = subs[0];

                            field2 = subs[1];
                            field3 = subs[2];
                            field4 = subs[3];
                            field5 = subs[4];
                            field6 = subs[5];
                            if ((field6[0] == '@') || (field6[0] == '.')) field6 = "foo";
                            if (field5 == "00000000") field6 = "foo";
                        }
                    }
                    else
                    {
                        field6 = "foo";
                    }
                }
                else
                {
                    field6 = "foo";
                }
            }
            else
            {
                field6 = "foo";
            }
        }
    }

class Symbol
    {
        internal string name;
        internal string address;
        internal string section;
        internal int dimension;

        public Symbol ()
        {
            name = "foo";
            address = "00000000";
        }

        public Symbol(ref RawSymbol R)
        {

            name = R.field6;
            if (name != "foo")
            {
                address = R.field1;
                section = R.field4;
                try
                { 
                    dimension = Convert.ToInt16(R.field5, 16);
                }
                catch
                {
                    dimension = 0;
                }
            }
        }

        internal string ConvertToLine()
        {
            if (this.name != "foo")
            {
                return (this.name + "\t" + this.address + "\t" + Convert.ToString(dimension) + "\t" + this.section);
            }
            else
            { 
                return (null);
            }
        }
    }
}
