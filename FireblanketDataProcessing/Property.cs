using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FireblanketDataProcessing
{
    public class Property : IComparable
    {
        public string BoroughNumber { get; set; }
        public string BoroughName { get; set; }
        public string Block { get; set; }
        public string Lot { get; set; }
        public string PARID { get; set; }
        public string ZIP { get; set; }
        public string Address { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string Neighborhood { get; set; }
        public List<Tuple<string, string>> YearsAndValues = new List<Tuple<string, string>>();
        public bool Easement { get; set; }

        //Temporary values
        public string Value { get; set; }
        public string Year { get; set; }

        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }

            Property p = obj as Property;

            if (p != null)
            {
                return PARID.CompareTo(p.PARID);
            }
            else
            {
                throw new Exception();
            }
        }
    }
}
