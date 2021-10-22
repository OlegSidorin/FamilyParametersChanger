using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilyParametersChanger
{
    public enum Causes
    {
        UnknownState = 0,
        WrongGuidAndName = 1,
        GoodNameWrongGuid = 2,
        GoodGuidWrongName = 3,
        OK = 4
    }
}
