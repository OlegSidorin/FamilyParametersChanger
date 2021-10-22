using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System.Windows;
using System.Linq;

namespace FamilyParametersChanger
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class MainCommand1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            if (!doc.IsFamilyDocument) return Result.Succeeded;

            #region Repeat TEST part
            FilePathViewModel filePathViewModel = new FilePathViewModel();
            filePathViewModel.FileType = FileType.Family;
            filePathViewModel.PathString = doc.PathName;
            filePathViewModel.Name = doc.Title + ".rfa";

            List<SharedParameterElement> sharedParameters = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();

            List<ParameterAndFamily> listOfParametersAndComments = new List<ParameterAndFamily>();
            foreach (SharedParameterElement parameter in sharedParameters)
            {
                Report report;
                string guid = parameter.GuidValue.ToString();
                string name = parameter.Name;
                if (FOP.IsSomethingWrongWithParameter(guid, name, out report))
                {
                    var parameterAndFamily = new ParameterAndFamily()
                    {
                        ParameterName = name,
                        ParameterGuid = guid,
                        Cause = report.Cause,
                        Comment = report.Comment
                    };
                    listOfParametersAndComments.Add(parameterAndFamily);
                }
            }

            var familySymbols = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).ToElements();
            var families = new List<Family>();
            foreach (FamilySymbol fs in familySymbols)
            {
                bool familyAlreadyIs = false;
                foreach (Family f in families)
                {
                    if (f.Name == fs.FamilyName)
                        familyAlreadyIs = true;
                }
                if (!familyAlreadyIs)
                    families.Add(fs.Family);
            }
            List<Document> familyDocs = new List<Document>();
            var listOfParametersAndFamilies = new List<ParameterAndFamily>();
            foreach (Family f in families)
            {
                foreach (var shPar in listOfParametersAndComments)
                {

                    try
                    {
                        if (f.IsEditable)
                        {
                            Document familyDoc = doc.EditFamily(f);
                            if (null != familyDoc && familyDoc.IsFamilyDocument == true)
                            {
                                List<SharedParameterElement> familySharedParams = new FilteredElementCollector(familyDoc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();
                                foreach (var sp in familySharedParams)
                                {
                                    if (sp.GetDefinition().Name == shPar.ParameterName)
                                    {
                                        listOfParametersAndFamilies.Add(new ParameterAndFamily()
                                        {
                                            ParameterName = shPar.ParameterName,
                                            ParameterGuid = shPar.ParameterGuid,
                                            Cause = shPar.Cause,
                                            FamilyName = f.Name
                                        });
                                    }
                                }
                                familyDocs.Add(familyDoc);
                                //familyDoc.Close(false);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            foreach (Document familyDoc in familyDocs)
            {
                try
                {
                    familyDoc.Close(false);
                }
                catch { }
            }

            ParameterAndFamily PF = new ParameterAndFamily();
            var listOfParametersAndFamiliesDistinct = PF.GetDistinct(listOfParametersAndFamilies);
            var listOfParameterAndFamilies = PF.GetParametersWithListOfFamilies(listOfParametersAndFamiliesDistinct);
            var listOfFamilyAndParameters = PF.GetFamilyWithListOfParameters(listOfParametersAndFamiliesDistinct);

            var listOfParametersWithOutFamilies = new List<ParameterAndFamily>();
            foreach (var p in listOfParametersAndComments)
            {
                bool isin = false;
                foreach (var pf in listOfParameterAndFamilies)
                {
                    if (p.ParameterName == pf.ParameterName)
                    {
                        isin = true;
                    }
                }
                if (!isin)
                {
                    listOfParametersWithOutFamilies.Add(p);
                }
            }
            #endregion
            // Перевод параметров из общих в семейные
            string outputReport = "----------- О Т Ч Е Т -----------";

            List<WrongParameter> wrongParametersList = new List<WrongParameter>();
            foreach (ParameterAndFamily pf in listOfParameterAndFamilies)
            {
                if (pf.Cause == Causes.GoodGuidWrongName)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GuidValue.ToString() == pf.ParameterGuid)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.GoodGuidWrongName,
                                ParameterAndFamily = pf,
                                Families = new List<Family>()
                            };
                            foreach (string fn in pf.FamilyNames)
                            {
                                foreach (FamilySymbol fs in familySymbols)
                                {
                                    if (fs.FamilyName == fn)
                                    {
                                        bool isFamilyInList = false;
                                        foreach (Family family in wrongParameter.Families)
                                        {
                                            if (family.Name == fs.FamilyName)
                                                isFamilyInList = true;
                                        }
                                        if (!isFamilyInList)
                                            wrongParameter.Families.Add(fs.Family);
                                    }
                                }
                            }
                            wrongParametersList.Add(wrongParameter);
                        }
                    }
                }
                if (pf.Cause == Causes.WrongGuidAndName)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GetDefinition().Name == pf.ParameterName)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.WrongGuidAndName,
                                ParameterAndFamily = pf,
                                Families = new List<Family>()
                            };
                            wrongParametersList.Add(wrongParameter);
                        }
                    }
                }
            }

            List<WrongParameter> wrongParametersListWithOutFamilies = new List<WrongParameter>();
            foreach (ParameterAndFamily pf in listOfParametersWithOutFamilies)
            {
                if (pf.Cause == Causes.GoodGuidWrongName)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GuidValue.ToString() == pf.ParameterGuid)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.GoodGuidWrongName,
                                ParameterAndFamily = pf,
                            };
                            wrongParametersListWithOutFamilies.Add(wrongParameter);
                        }
                    }
                }
                if (pf.Cause == Causes.WrongGuidAndName)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GetDefinition().Name == pf.ParameterName)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.WrongGuidAndName,
                                ParameterAndFamily = pf,
                            };
                            wrongParametersListWithOutFamilies.Add(wrongParameter);
                        }
                    }
                }
            }

            foreach (WrongParameter wp in wrongParametersListWithOutFamilies)
            {
                switch (wp.Cause)
                {
                    case Causes.WrongGuidAndName:
                        outputReport += ConvertSharedParameterInFamily(doc, wp);
                        break;
                }
            }
            if (outputReport == "") outputReport = "Ничего не сделано!";
            MessageBox.Show(outputReport);

            return Result.Succeeded;
        }
        private string ConvertSharedParameterInFamily(Document doc, WrongParameter wp)
        {
            string outputReport = "";
            using (Transaction t1 = new Transaction(doc, "Convert Shared Parameter to Family Parameter With New Name"))
            {
                t1.Start();
                Definition d = wp.WrongSharedParameterElement.GetDefinition();
                InternalDefinition id = d as InternalDefinition;
                try
                {
                    FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.WrongSharedParameterElement.GetDefinition().Name).FirstOrDefault();
                    //FamilyParameter newfp = familyDoc.FamilyManager.AddParameter(fp.Definition.Name, fp.Definition.ParameterGroup, fp.Definition.ParameterType, fp.IsInstance);
                    doc.FamilyManager.ReplaceParameter(fp, fp.Definition.Name + "_new", wp.WrongSharedParameterElement.GetDefinition().ParameterGroup, fp.IsInstance);
                    //System.Windows.Forms.MessageBox.Show("1. замена общего параметра " + fp.Definition.Name + " на параметр семейства в семействе ( + new)" + fwp.Family.Name);
                    outputReport += "\nЗамена общего параметра " + fp.Definition.Name + " на параметр семейства";
                }
                catch (Exception ex)
                {
                    outputReport += "\nНе удалось заменить на параметр семейства (возможно, он во вложенном семействе) " + wp.ParameterAndFamily.ParameterName + " на параметр семейства";
                };
                t1.Commit();
            }
            using (Transaction t2 = new Transaction(doc, "Rename Parameter On Old Names"))
            {
                t2.Start();
                List<FamilyParameter> fpset = doc.FamilyManager.GetParameters().ToList();
                //var fps = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilyParameter)).Cast<FamilyParameter>().ToList();
                //str += "\n" + fi.Name + " " + family.Name + "\n";
                foreach (var fp in fpset)
                {
                    if (fp.Definition.Name.Contains("_new"))
                    {
                        try
                        {
                            string newname = fp.Definition.Name.Replace("_new", "");
                            doc.FamilyManager.RenameParameter(fp, newname);
                            //System.Windows.Forms.MessageBox.Show("2. возврат к прежнему имени (- new) параметра" + fp.Definition.Name + " в семействе " + fwp.Family.Name);
                            //outputReport += "\nВозврат к прежнему имени (- _new) параметра " + fp.Definition.Name + " в семействе " + doc.Title;
                        }
                        catch (Exception ex)
                        {
                            outputReport += "\nНе удалось вернуть преженее имя " + wp.ParameterAndFamily.ParameterName + " в семействе " + doc.Title;
                        }

                    }

                    //str += "    fp->>" + fp.Definition.Name + "\n";
                }
                t2.Commit();
            }
            using (Transaction t3 = new Transaction(doc, "Delete Parameter"))
            {
                t3.Start();
                var sps = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();
                foreach (SharedParameterElement sharedParameterElement in sps)
                {
                    if (sharedParameterElement.GetDefinition().Name == wp.ParameterAndFamily.ParameterName)
                    {
                        try
                        {
                            string result = DeleteSP(sharedParameterElement, doc);
                            //System.Windows.Forms.MessageBox.Show("3. удаление общего параметра " + result + " в семействе " + fwp.Family.Name);
                            outputReport += "\nУдаление общего параметра " + wp.ParameterAndFamily.ParameterName + ": " + result;
                        }
                        catch (Exception ex)
                        {
                            outputReport += "\nНе удалось удалить общий параметр " + wp.ParameterAndFamily.ParameterName + " в семействе " + doc.Title;
                        }
                    }
                }
                t3.Commit();
            }
            return outputReport;
        }
        public string DeleteSP(SharedParameterElement sp, Document doc)
        {
            string output = "\nпараметр в семействе не обнаружен ";
            if (sp != null)
            {
                try
                {
                    output = "\nпараметр " + sp.GetDefinition().Name;
                    var outputid = doc.Delete(sp.GetDefinition().Id);
                    string t = "";
                    foreach (var item in outputid)
                    {
                        t += item.IntegerValue;
                    }
                    if (t != null)
                    {
                        return output + " успешно удален";
                    }
                    return t;
                }
                catch (Exception ex)
                {
                    return output + "и не удален";
                }
            }
            return output;
        }
    }
}
