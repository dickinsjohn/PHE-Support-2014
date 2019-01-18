using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PHE_Supports
{
    //structure to store specification data from file
    public struct SpecificationData
    {
        public string selectedFamily;
        public string discipline;
        public double offset;
        public double minSpacing;
        public string supportType;
        public string specsFile;
    };
}
