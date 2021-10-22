using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilyParametersChanger
{
    public static class Extensions
    {
        public static string ToFriendlyString(this Causes cause)
        {
            switch (cause)
            {
                case Causes.OK:
                    return "с параметром все ок";
                case Causes.WrongGuidAndName:
                    return "параметр не из ФОП";
                case Causes.GoodGuidWrongName:
                    return "guid ок, имя нет";
                case Causes.GoodNameWrongGuid:
                    return "имя ок, guid не верен";
                case Causes.UnknownState:
                    return "не понятно";
                default:
                    return "?";
            }
        }
        public static string ToFriendlyString(this FileType fileType)
        {
            switch (fileType)
            {
                case FileType.Folder:
                    return "Folder";
                case FileType.Family:
                    return "Family";
                case FileType.Project:
                    return "Project";
                case FileType.Template:
                    return "Template";
                case FileType.Unknown:
                    return "Не понятно";
                default:
                    return "?";
            }
        }
    }
}
