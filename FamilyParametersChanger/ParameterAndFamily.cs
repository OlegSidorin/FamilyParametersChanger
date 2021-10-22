using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilyParametersChanger
{
    public class ParameterAndFamily
    {
        public string ParameterName { get; set; }
        public string ParameterGuid { get; set; }
        public string Comment { get; set; }
        public string FamilyName { get; set; }
        public List<string> FamilyNames { get; set; }
        public List<string> ParameterNames { get; set; }
        public Causes Cause { get; set; }

        public List<ParameterAndFamily> GetDistinct(List<ParameterAndFamily> inputList)
        {
            var outputList = new List<ParameterAndFamily>();
            bool isInList(ParameterAndFamily pf, List<ParameterAndFamily> listOfPF)
            {
                foreach (var item in listOfPF)
                {
                    if (item.ParameterName == pf.ParameterName)
                    {
                        if (item.FamilyName == pf.FamilyName)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }

            foreach (ParameterAndFamily pf in inputList)
            {
                if (!isInList(pf, outputList))
                {
                    outputList.Add(pf);
                }
            }
            return outputList;
        }

        public List<ParameterAndFamily> GetParametersWithListOfFamilies(List<ParameterAndFamily> inputList)
        {
            var outputList = new List<ParameterAndFamily>();

            bool parameterNameIsInList(ParameterAndFamily PF, List<ParameterAndFamily> PFList)
            {
                foreach (var pf in PFList)
                {
                    if ((pf.ParameterName == PF.ParameterName))
                    {
                        if (pf.ParameterGuid == PF.ParameterGuid)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }

            ParameterAndFamily GetElementWithSameParameter(ParameterAndFamily PF, List<ParameterAndFamily> parameterAndFamilies)
            {
                foreach (var pf in parameterAndFamilies)
                {
                    if (PF.ParameterName == pf.ParameterName)
                    {
                        if (PF.ParameterGuid == pf.ParameterGuid)
                        {
                            return pf;
                        }
                    }
                }
                return new ParameterAndFamily();
            }

            foreach (var pf in inputList)
            {
                if (!parameterNameIsInList(pf, outputList))
                {
                    ParameterAndFamily newPF = new ParameterAndFamily()
                    {
                        ParameterName = pf.ParameterName,
                        ParameterGuid = pf.ParameterGuid,
                        Cause = pf.Cause,
                        FamilyNames = new List<string>()
                    };
                    newPF.FamilyNames.Add(pf.FamilyName);
                    outputList.Add(newPF);
                }
                else
                {
                    GetElementWithSameParameter(pf, outputList).FamilyNames.Add(pf.FamilyName);
                }
            }

            return outputList;
        }

        public static ParameterAndFamily GetParameterUsingName(string name, List<ParameterAndFamily> parameterAndFamilies)
        {
            foreach (ParameterAndFamily parameterAndFamily in parameterAndFamilies)
            {
                if (parameterAndFamily.ParameterName == name)
                {
                    return parameterAndFamily;
                }
            }
            return new ParameterAndFamily();
        }

        public string GetFamiliesInOneSting()
        {
            string output = "";
            foreach (string s in FamilyNames)
            {
                output += s + ", ";
            }
            return output.Remove(output.Length - 2, 2);
        }

        public List<ParameterAndFamily> GetFamilyWithListOfParameters(List<ParameterAndFamily> inputList)
        {
            var outputList = new List<ParameterAndFamily>();

            bool familyNameIsInList(ParameterAndFamily PF, List<ParameterAndFamily> PFList)
            {
                foreach (var pf in PFList)
                {
                    if ((pf.FamilyName == PF.FamilyName))
                    {
                        return true;
                    }
                }
                return false;
            }

            ParameterAndFamily GetElementWithSameFamily(ParameterAndFamily PF, List<ParameterAndFamily> parameterAndFamilies)
            {
                foreach (var pf in parameterAndFamilies)
                {
                    if (PF.FamilyName == pf.FamilyName)
                    {
                        return pf;
                    }
                }
                return new ParameterAndFamily();
            }

            foreach (var pf in inputList)
            {
                if (!familyNameIsInList(pf, outputList))
                {
                    ParameterAndFamily newPF = new ParameterAndFamily()
                    {
                        FamilyName = pf.FamilyName,
                        ParameterNames = new List<string>()
                    };
                    newPF.ParameterNames.Add(pf.ParameterName); //(pf.ParameterName + $" ( {pf.ParameterGuid} : {pf.Cause.ToFriendlyString()} )");
                    outputList.Add(newPF);
                }
                else
                {
                    GetElementWithSameFamily(pf, outputList).ParameterNames.Add(pf.ParameterName); //(pf.ParameterName + $" ( {pf.ParameterGuid} : {pf.Cause.ToFriendlyString()} )");
                }
            }

            return outputList;
        }

        public List<ParameterAndFamily> GetParametersForDeleting(List<ParameterAndFamily> input)
        {
            List<ParameterAndFamily> output = new List<ParameterAndFamily>();
            foreach (ParameterAndFamily pf in input)
            {
                if (pf.FamilyNames.Count <= 1)
                {
                    output.Add(pf);
                }
            }
            return output;
        }

        public string GetParametersInOneSting()
        {
            string output = "";
            foreach (string s in ParameterNames)
            {
                output += s + ", ";
            }
            return output.Remove(output.Length - 2, 2);
        }
    }
}
