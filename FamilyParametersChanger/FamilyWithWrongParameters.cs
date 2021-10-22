using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilyParametersChanger
{
    public class FamilyWithWrongParameters
    {
        public Family Family { get; set; }
        public List<WrongParameter> WrongParameters { get; set; }
    }
}
