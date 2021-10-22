using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilyParametersChanger
{
    public class WrongParameter
    {
        public SharedParameterElement WrongSharedParameterElement { get; set; }
        public Causes Cause { get; set; }
        public List<Family> Families { get; set; }
        public ParameterAndFamily ParameterAndFamily { get; set; }
    }
}
