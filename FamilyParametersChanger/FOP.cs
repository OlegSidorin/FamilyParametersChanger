using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FamilyParametersChanger
{
    public static class FOP
    {
        public static string PathToFOP { get; } = Main.DllFolderLocation + @"\res\ФОП2019.txt";
        public static List<SharedParameterFromFOP> SharedParametersFromFOP { get; } = GetSharedParametersFromFOP();

        public static SharedParameterFromFOP GetParameterFromFOPUsingGUID(string guid)
        {
            List<SharedParameterFromFOP> parameters = GetSharedParametersFromFOP();
            foreach (var p in parameters)
            {
                if (p.Guid == guid)
                    return p;
            }
            return new SharedParameterFromFOP()
            {
                Name = "",
                Guid = "",
                Group = "",
                Type = ""
            };
        }
        public static SharedParameterFromFOP GetParameterFromFOPUsingName(string name)
        {
            List<SharedParameterFromFOP> parameters = GetSharedParametersFromFOP();
            foreach (var p in parameters)
            {
                if (p.Name == name)
                    return p;
            }
            return new SharedParameterFromFOP()
            {
                Name = "",
                Guid = "",
                Group = "",
                Type = ""
            };
        }

        public static bool IsSomethingWrongWithParameter(string guid, string name, out Report report)
        {
            bool output = true;
            foreach (var par in SharedParametersFromFOP)
            {
                if (guid == par.Guid && name == par.Name)
                {
                    report = new Report()
                    {
                        Cause = Causes.OK,
                        Comment = "Параметр из ФОП"
                    };
                    return false;
                }
                if (guid != par.Guid && name == par.Name)
                {
                    report = new Report()
                    {
                        Cause = Causes.GoodNameWrongGuid,
                        Comment = "Параметр имеет неверный guid"
                    };
                    return true;
                }
                if (guid == par.Guid && name != par.Name)
                {
                    report = new Report()
                    {
                        Cause = Causes.GoodGuidWrongName,
                        Comment = "Параметр имеет неверное имя"
                    };
                    return true;
                }
                //if (guid != par.Guid && name != par.Name)
                //{
                //    report = new Report()
                //    {
                //        Cause = Causes.UnknownState,
                //        Comment = ""
                //    };
                //}
            }
            report = new Report()
            {
                Cause = Causes.WrongGuidAndName,
                Comment = "Параметр не из ФОП"
            };
            return output;
        }

        public static List<SharedParameterFromFOP> GetSharedParametersFromFOP()
        {
            List<SharedParameterFromFOP> outputList = new List<SharedParameterFromFOP>();
            string line;
            string guid, name, type, ctype, group;

            int startGuid, endGuid, startName, endName, startType, endType, startCtype, endCtype, startGroup, endGroup;

            List<GroupFromFop> gFOP = GetGroups();
            string getGroup(List<GroupFromFop> groups, string groupNumber)
            {

                foreach (var gr in groups)
                {
                    if (gr.Number == groupNumber)
                        return gr.Name;
                }
                return "";
            }

            // Read the file line by line.  
            System.IO.StreamReader file = new System.IO.StreamReader(PathToFOP);
            while ((line = file.ReadLine()) != null)
            {

                if (line.Contains("PARAM") && !line.Contains("*"))
                {
                    startGuid = line.IndexOf('\t', 0) + 1;
                    endGuid = line.IndexOf('\t', startGuid + 1);
                    guid = line.Substring(startGuid, endGuid - startGuid);
                    startName = endGuid + 1;
                    endName = line.IndexOf('\t', startName + 1);
                    name = line.Substring(startName, endName - startName);
                    startType = endName + 1;
                    endType = line.IndexOf('\t', startType + 1);
                    type = line.Substring(startType, endType - startType);
                    startCtype = endType + 1;
                    endCtype = line.IndexOf('\t', startCtype);
                    ctype = line.Substring(startCtype, endCtype - startCtype);
                    startGroup = endCtype + 1;
                    endGroup = line.IndexOf('\t', startGroup);
                    group = line.Substring(startGroup, endGroup - startGroup);
                    SharedParameterFromFOP sharedParameterFromFOP = new SharedParameterFromFOP()
                    {
                        Guid = guid,
                        Name = name,
                        Type = type,
                        Group = getGroup(gFOP, group)
                    };
                    outputList.Add(sharedParameterFromFOP);
                }
            }

            file.Close();
            return outputList;
        }

        public static List<GroupFromFop> GetGroups()
        {
            string line;

            string id, gname;
            int startId, endId, startN, endN;

            List<GroupFromFop> outputList = new List<GroupFromFop>();
            System.IO.StreamReader file = new System.IO.StreamReader(PathToFOP);
            while ((line = file.ReadLine()) != null)
            {

                if (line.Contains("GROUP") && !line.Contains("*"))
                {
                    startId = line.IndexOf('\t', 0) + 1;
                    endId = line.IndexOf('\t', startId);
                    id = line.Substring(startId, endId - startId);
                    startN = endId + 1;
                    gname = "";
                    try
                    {
                        gname = line.Substring(startN, line.Length - startN);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                    }


                    GroupFromFop groupFromFOP = new GroupFromFop()
                    {
                        Number = id,
                        Name = gname
                    };
                    outputList.Add(groupFromFOP);
                }
            }

            file.Close();
            return outputList;
        }

    }
}
