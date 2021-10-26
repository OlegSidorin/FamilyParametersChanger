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
using System.Linq;
using System.Windows.Forms;

namespace FamilyParametersChanger
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class MainCommand0 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            commandData.Application.Application.SharedParametersFilename = FOP.PathToFOP;
            if (!doc.IsFamilyDocument) return Result.Succeeded;

            #region Repeat TEST part
            FilePathViewModel filePathViewModel = new FilePathViewModel();
            filePathViewModel.FileType = FileType.Family;
            filePathViewModel.PathString = doc.PathName;
            filePathViewModel.Name = doc.Title + ".rfa";

            List<SharedParameterElement> sharedParameters = new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>().ToList();

            // формирование списка параметров с комментариями
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
            //string st23 = "";
            //foreach (ParameterAndFamily pff in listOfFamilyAndParameters)
            //{
            //    st23 += "\n" + pff.FamilyName + "\n";
            //    foreach (var it in pff.ParameterNames)
            //    {
            //        st23 += it;
            //    }
            //}
            //MessageBox.Show(st23);
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

            string outputReport = "";
            FlowDocumentForReport flowDocumentForReport = new FlowDocumentForReport();
            flowDocumentForReport.FlowDocument = new System.Windows.Documents.FlowDocument();
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
                if (pf.Cause == Causes.GoodNameWrongGuid)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GetDefinition().Name == pf.ParameterName)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.GoodNameWrongGuid,
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
                if (pf.Cause == Causes.GoodNameWrongGuid)
                {
                    foreach (SharedParameterElement sp in sharedParameters)
                    {
                        if (sp.GetDefinition().Name == pf.ParameterName)
                        {
                            WrongParameter wrongParameter = new WrongParameter()
                            {
                                WrongSharedParameterElement = sp,
                                Cause = Causes.GoodNameWrongGuid,
                                ParameterAndFamily = pf,
                            };
                            wrongParametersListWithOutFamilies.Add(wrongParameter);
                        }
                    }
                }
            }

            List<FamilyWithWrongParameters> familiesWithWrongParameters = new List<FamilyWithWrongParameters>();
            foreach (ParameterAndFamily fl in listOfFamilyAndParameters)
            {
                foreach (FamilySymbol fs in familySymbols)
                {
                    if (fs.FamilyName == fl.FamilyName)
                    {
                        FamilyWithWrongParameters familyWithWrongParameters = new FamilyWithWrongParameters()
                        {
                            Family = fs.Family,
                            WrongParameters = new List<WrongParameter>()
                        };
                        bool familyIsIn = false;
                        foreach (var item in familiesWithWrongParameters)
                        {
                            if (item.Family.Name == fs.FamilyName)
                                familyIsIn = true;
                        }
                        if (!familyIsIn)
                        {
                            foreach (string parameterName in fl.ParameterNames)
                            {
                                foreach (WrongParameter wrongParameter in wrongParametersList)
                                {
                                    bool paramIsIn = false;
                                    foreach(var item in familyWithWrongParameters.WrongParameters)
                                    {
                                        if (item.ParameterAndFamily.ParameterName == parameterName)
                                        {
                                            paramIsIn = true;
                                        }
                                    }
                                    if (parameterName == wrongParameter.ParameterAndFamily.ParameterName)
                                    {
                                        if (!paramIsIn)
                                        {
                                            familyWithWrongParameters.WrongParameters.Add(wrongParameter);
                                        }
                                    }
                                }
                            }
                            familiesWithWrongParameters.Add(familyWithWrongParameters);
                        }
                    }
                }
            }
            //string srt123 = "";
            //foreach (var item1 in familiesWithWrongParameters)
            //{
            //    srt123 += "\n" + item1.Family.Name + "\n";
            //    foreach(var item2 in item1.WrongParameters)
            //    {
            //        srt123 += item2.WrongSharedParameterElement.GetDefinition().Name + " " + item2.Cause.ToFriendlyString() + "\n";
                    
            //    }
            //}
            
            //MessageBox.Show(srt123);
            //return Result.Succeeded;
            // Action
            int repeat = 0;
            
            foreach (WrongParameter wp in wrongParametersListWithOutFamilies)
            {
                if (repeat == 0)
                {
                    flowDocumentForReport.AddHead("Исправление параметров, которых нет во вложенных семействах:");
                    repeat = 1;
                }
                switch (wp.Cause)
                {
                    case Causes.WrongGuidAndName:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success, out string str11, out string str12, out string str13);
                        flowDocumentForReport.AddParagraph(str11, str12, str13);
                        if(success)
                        {
                            outputReport += DeleteWP(doc, wp, out bool success1, out string str14, out string str15, out string str16);
                            flowDocumentForReport.AddParagraph(str14, str15, str16);
                        }
                        break;
                    case Causes.GoodNameWrongGuid:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success5, out string str21, out string str22, out string str23);
                        flowDocumentForReport.AddParagraph(str21, str22, str23);
                        if (success5)
                        {
                            outputReport += DeleteWP(doc, wp, out bool success6, out string str24, out string str25, out string str26);
                            flowDocumentForReport.AddParagraph(str24, str25, str26);
                        }
                        break;
                    case Causes.GoodGuidWrongName:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success2, out string str31, out string str32, out string str33);
                        flowDocumentForReport.AddParagraph(str31, str32, str33);
                        if (success2)
                        {
                            outputReport += DeleteWP(doc, wp, out bool success3, out string str34, out string str35, out string str36);
                            flowDocumentForReport.AddParagraph(str34, str35, str36);
                            outputReport += ConvertFamilyParameterInToSharedParameter_NAMEFIX(doc, wp, commandData, out bool success4, out string str37, out string str38, out string str39);
                            flowDocumentForReport.AddParagraph(str37, str38, str39);
                        }
                        break;
                }
            }

            List<string> pathsToTempFamilies = new List<string>();
            List<WrongParameter> wrongParametersListThatInFamilies = new List<WrongParameter>();
            repeat = 0;
            foreach (FamilyWithWrongParameters fwp in familiesWithWrongParameters)
            {
                if (repeat == 0)
                {
                    flowDocumentForReport.AddHead("Исправление параметров во вложенных семействах:");
                    repeat = 1;
                }
                try
                {
                    Document familyDoc = doc.EditFamily(fwp.Family);
                    outputReport += "\n\nОткрытие семейства " + familyDoc.Title;
                    flowDocumentForReport.AddParagraph("Открытие семейства:", familyDoc.Title, "");
                    if (null != familyDoc && familyDoc.IsFamilyDocument == true)
                    {
                        //string st = "";
                        //foreach (WrongParameter wp in fwp.WrongParameters)
                        //{
                        //    st += wp.ParameterAndFamily.ParameterName + "\n";
                        //}
                        //MessageBox.Show(st);
                        foreach (WrongParameter wp in fwp.WrongParameters)
                        {
                            switch (wp.Cause)
                            {
                                case Causes.WrongGuidAndName:
                                    outputReport += ConvertSharedParameterInToFamilyParameter(familyDoc, wp, out bool success, out string str11, out string str12, out string str13);
                                    flowDocumentForReport.AddParagraph(str11, str12, str13);
                                    if (success)
                                    {
                                        outputReport += DeleteWP(familyDoc, wp, out bool success1, out string str14, out string str15, out string str16);
                                        flowDocumentForReport.AddParagraph(str14, str15, str16);
                                        wrongParametersListThatInFamilies.Add(wp);
                                    }
                                    break;
                                case Causes.GoodNameWrongGuid:
                                    outputReport += ConvertSharedParameterInToFamilyParameter(familyDoc, wp, out bool success5, out string str21, out string str22, out string str23);
                                    flowDocumentForReport.AddParagraph(str21, str22, str23);
                                    if (success5)
                                    {
                                        outputReport += DeleteWP(familyDoc, wp, out bool success6, out string str24, out string str25, out string str26);
                                        flowDocumentForReport.AddParagraph(str24, str25, str26);
                                        outputReport += ConvertFamilyParameterInToSharedParameter_GUIDFIX(familyDoc, wp, commandData, out bool success7, out string str27, out string str28, out string str29);
                                        flowDocumentForReport.AddParagraph(str27, str28, str29);
                                        wrongParametersListThatInFamilies.Add(wp);
                                    }
                                    break;
                                case Causes.GoodGuidWrongName:
                                    outputReport += ConvertSharedParameterInToFamilyParameter(familyDoc, wp, out bool success2, out string str31, out string str32, out string str33);
                                    flowDocumentForReport.AddParagraph(str31, str32, str33);
                                    if (success2)
                                    {
                                        outputReport += DeleteWP(familyDoc, wp, out bool success3, out string str34, out string str35, out string str36);
                                        flowDocumentForReport.AddParagraph(str34, str35, str36);
                                        outputReport += ConvertFamilyParameterInToSharedParameter_NAMEFIX(familyDoc, wp, commandData, out bool success4, out string str37, out string str38, out string str39);
                                        flowDocumentForReport.AddParagraph(str37, str38, str39);
                                        wrongParametersListThatInFamilies.Add(wp);
                                    }
                                    break;
                            }
                        }
                        string tmpFile = Path.Combine(@"C:\Temp", fwp.Family.Name + ".rfa");
                        if (File.Exists(tmpFile))
                            File.Delete(tmpFile);
                        familyDoc.SaveAs(tmpFile);
                        pathsToTempFamilies.Add(tmpFile);
                        //System.Windows.Forms.MessageBox.Show("Конец редактирования семейства: " + family.Name);
                        outputReport += "\nЗавершение редактирования семейства: " + familyDoc.Title;
                        flowDocumentForReport.AddParagraph("Завершение редактирования семейства:", familyDoc.Title, "");
                        familyDoc.Close(false);
                    }
                    else
                    {
                        familyDoc.Close(false);
                        outputReport += "\nЗакрытие семейства " + familyDoc.Title;
                        flowDocumentForReport.AddParagraph("Аварийное закрытие семейства:", familyDoc.Title, ";");
                    }
                }
                catch (Exception ex)
                {
                    outputReport += "\nНе удалось открыть для редактирования";
                    flowDocumentForReport.AddParagraph("Не удалось открыть для редактирования", "", "");
                }
            }

            //outputReport += "\n\nУдаление общих нежелательных параметров: ";
            repeat = 0;
            foreach (WrongParameter wp in wrongParametersListThatInFamilies)
            {
                if (repeat == 0)
                {
                    flowDocumentForReport.AddHead("Удаление нежелательных общих параметров из семейства, которые находятся во вложенных семействах:");
                    repeat = 1;
                }
                switch (wp.Cause)
                {
                    case Causes.WrongGuidAndName:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success, out string str101, out string str102, out string str103);
                        if (success) flowDocumentForReport.AddParagraph(str101, str102, str103);
                        outputReport += DeleteWP(doc, wp, out bool success1, out string str104, out string str105, out string str106);
                        if (success1) flowDocumentForReport.AddParagraph(str104, str105, str106);
                        break;
                    case Causes.GoodNameWrongGuid:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success5, out string str201, out string str202, out string str203);
                        if (success5) flowDocumentForReport.AddParagraph(str201, str202, str203);
                        outputReport += DeleteWP(doc, wp, out bool success6, out string str204, out string str205, out string str206);
                        if (success6) flowDocumentForReport.AddParagraph(str204, str205, str206);
                        if (success5 && success6)
                        {
                            outputReport += ConvertFamilyParameterInToSharedParameter_GUIDFIX(doc, wp, commandData, out bool success7, out string str207, out string str208, out string str209);
                            if (success7) flowDocumentForReport.AddParagraph(str207, str208, str209);
                        }
                        break;
                    case Causes.GoodGuidWrongName:
                        outputReport += ConvertSharedParameterInToFamilyParameter(doc, wp, out bool success2, out string str301, out string str302, out string str303);
                        if (success2) flowDocumentForReport.AddParagraph(str301, str302, str303);
                        outputReport += DeleteWP(doc, wp, out bool success3, out string str304, out string str305, out string str306);
                        if (success3) flowDocumentForReport.AddParagraph(str304, str305, str306);
                        if (success2 && success3)
                        { 
                            outputReport += ConvertFamilyParameterInToSharedParameter_NAMEFIX(doc, wp, commandData, out bool success4, out string str307, out string str308, out string str309);
                            if (success4) flowDocumentForReport.AddParagraph(str307, str308, str309);
                        }
                        break;
                }

                    //outputReport += DeleteWP(doc, wp, out bool succcess, out string str11, out string str12, out string str13);
                    //if (str11 != "") flowDocumentForReport.AddParagraph(str11, str12, str13);

            }

            //outputReport += "\n\nЗагрузка семейств: ";
            repeat = 0;
            foreach (string path in pathsToTempFamilies)
            {
                if (repeat == 0)
                {
                    flowDocumentForReport.AddHead("Загрузка семейств:"); ;
                    repeat = 1;
                }
                using (Transaction trans = new Transaction(doc, "Load family."))
                {
                    trans.Start();
                    //System.Windows.Forms.MessageBox.Show("Начало загрузки семейства: " + Path.GetFileName(path) + " ");
                    try
                    {
                        IFamilyLoadOptions famLoadOptions = new JtFamilyLoadOptions();
                        Family newFam = null;
                        bool yes = doc.LoadFamily(path, new JtFamilyLoadOptions(), out newFam);
                        if (yes)
                            File.Delete(path);
                        //System.Windows.Forms.MessageBox.Show("7. Загружено семейство " + yes.ToString() + " " + Path.GetFileName(path) + " в " + doc.Title + ".rvt\n");
                        outputReport += "\nЗагружено семейство " + Path.GetFileName(path) + " (" + yes + ") ";
                        flowDocumentForReport.AddParagraph("Загружено семейство", Path.GetFileName(path), "");
                    }
                    catch (Exception ex)
                    {
                        //System.Windows.Forms.MessageBox.Show(ex.ToString());
                        outputReport += "\nНе удалось загрузить семейство " + Path.GetFileName(path);
                        flowDocumentForReport.AddParagraph("Не удалось загрузить семейство", Path.GetFileName(path), "");
                    }

                    trans.Commit();
                    //System.Windows.Forms.MessageBox.Show("Конец загрузки семейства: " + Path.GetFileName(path) + " ");

                }
            }


            //if (outputReport == "") outputReport = "Ничего не сделано!";
            if (outputReport == "")
            {
                flowDocumentForReport.AddHead("Ничего не сделано");
            }
            //MessageBox.Show(outputReport);
            ReportWindow reportWindow = new ReportWindow();
            reportWindow.flowDocScrollViewer.Document = flowDocumentForReport.FlowDocument;
            reportWindow.Show();
            return Result.Succeeded;
        }
        private string ConvertSharedParameterInToFamilyParameter(Document doc, WrongParameter wp, out bool success, out string str1, out string pName, out string str2)
        {
            str1 = "";
            str2 = "";
            pName = "";
            success = true;
            string outputReport = "";
            using (Transaction t1 = new Transaction(doc, "Convert Shared Parameter to Family Parameter With New Name"))
            {
                t1.Start();
                try
                {
                    Definition d = wp.WrongSharedParameterElement.GetDefinition();
                    InternalDefinition id = d as InternalDefinition;
                    FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.ParameterAndFamily.ParameterName).FirstOrDefault();
                    //FamilyParameter newfp = familyDoc.FamilyManager.AddParameter(fp.Definition.Name, fp.Definition.ParameterGroup, fp.Definition.ParameterType, fp.IsInstance);
                    //MessageBox.Show(fp.Definition.Name + "\n" + wp.ParameterAndFamily.ParameterName + "\n" + wp.Cause.ToFriendlyString());
                    FamilyParameter newfp = doc.FamilyManager.ReplaceParameter(fp, wp.ParameterAndFamily.ParameterName + "_newfp", fp.Definition.ParameterGroup, fp.IsInstance); // <--------------------------------------
                    str1 = "Замена общего параметра";
                    pName = wp.ParameterAndFamily.ParameterName;
                    str2 = "на параметр семейства";
                    outputReport += $"\n{str1} {pName} {str2}";
                    //System.Windows.Forms.MessageBox.Show("1. замена общего параметра " + fp.Definition.Name + " на параметр семейства в семействе ( + new)" + fwp.Family.Name);

                }
                catch (Exception ex)
                {
                    success = false;
                    str1 = "Не удалось заменить";
                    pName = wp.ParameterAndFamily.ParameterName;
                    str2 = "на параметр семейства";
                    outputReport += $"\n{str1} {pName} {str2}";
                };
                t1.Commit();
            }
            if (success)
            {
                using (Transaction t2 = new Transaction(doc, "Rename Parameter On Old Names"))
                {
                    t2.Start();
                    try
                    {
                        FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.WrongSharedParameterElement.GetDefinition().Name + "_newfp").FirstOrDefault();
                        string newname = fp.Definition.Name.Replace("_newfp", "");
                        doc.FamilyManager.RenameParameter(fp, newname);
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        str1 = "Не удалось вернуть преженее имя";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = " в семействе";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                    t2.Commit();
                }
            }
            return outputReport;
        }
        private string ConvertFamilyParameterInToSharedParameter_NAMEFIX(Document doc, WrongParameter wp, ExternalCommandData commandData, out bool success, out string str1, out string pName, out string str2)
        {
            str1 = "";
            str2 = "";
            pName = "";
            success = true;
            string outputReport = "";
            using (Transaction t4 = new Transaction(doc, "Rename Parameters On New Name"))
            {
                t4.Start();
                List<FamilyParameter> fpset = doc.FamilyManager.GetParameters().ToList();
                foreach (var fp in fpset)
                {
                    try
                    {
                        if (fp.Definition.Name == wp.ParameterAndFamily.ParameterName)
                        {
                            doc.FamilyManager.RenameParameter(fp, fp.Definition.Name + "_newsp");
                        }
                    }
                    catch
                    {
                        success = false;
                        str1 = "Не удалось переименовать параметр семейства на новое временное имя";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                }
                t4.Commit();
            }
            if (success)
            {
                using (Transaction t5 = new Transaction(doc, "Replace Family Parameters With Shared Parameter"))
                {
                    t5.Start();
                    DefinitionFile definitionFile = commandData.Application.Application.OpenSharedParameterFile();
                    try
                    {
                        SharedParameterFromFOP sharedParameterFromFOP = FOP.GetParameterFromFOPUsingGUID(wp.ParameterAndFamily.ParameterGuid);
                        DefinitionGroup definitionGroup = definitionFile.Groups.get_Item(sharedParameterFromFOP.Group);
                        Definition definition = definitionGroup.Definitions.get_Item(sharedParameterFromFOP.Name);
                        ExternalDefinition externalDefinition = definition as ExternalDefinition;
                        FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.ParameterAndFamily.ParameterName + "_newsp").FirstOrDefault();
                        doc.FamilyManager.ReplaceParameter(fp, externalDefinition, fp.Definition.ParameterGroup, fp.IsInstance);
                        str1 = "Добавление общего параметра с правильным именем";
                        pName = definition.Name;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                    catch (Exception ex)
                    {
                        // здесь параметры, которые содержатся в семействах семейств нужно что то сделать с этим
                        //MessageBox.Show(ex.ToString());
                        success = false;
                        str1 = "Ошибка при замене на общий параметр";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                        //System.Windows.Forms.MessageBox.Show(wp.ParameterAndFamily.ParameterName + "\n" + ex.ToString());
                    };
                    t5.Commit();
                }
            }
            return outputReport;
        }
        private string ConvertFamilyParameterInToSharedParameter_GUIDFIX(Document doc, WrongParameter wp, ExternalCommandData commandData, out bool success, out string str1, out string pName, out string str2)
        {
            str1 = "";
            str2 = "";
            pName = "";
            success = true;
            string outputReport = "";
            using (Transaction t4 = new Transaction(doc, "Rename Parameters On New Name"))
            {
                t4.Start();
                List<FamilyParameter> fpset = doc.FamilyManager.GetParameters().ToList();
                foreach (var fp in fpset)
                {
                    try
                    {
                        if (fp.Definition.Name == wp.ParameterAndFamily.ParameterName)
                        {
                            doc.FamilyManager.RenameParameter(fp, fp.Definition.Name + "_newsp");
                        }
                    }
                    catch
                    {
                        success = false;
                        str1 = "Не удалось переименовать параметр семейства на новое временное имя";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                }
                t4.Commit();
            }
            if (success)
            {
                using (Transaction t5 = new Transaction(doc, "Replace Family Parameters With Shared Parameter"))
                {
                    t5.Start();
                    DefinitionFile definitionFile = commandData.Application.Application.OpenSharedParameterFile();
                    try
                    {
                        SharedParameterFromFOP sharedParameterFromFOP = FOP.GetParameterFromFOPUsingName(wp.ParameterAndFamily.ParameterName);
                        DefinitionGroup definitionGroup = definitionFile.Groups.get_Item(sharedParameterFromFOP.Group);
                        Definition definition = definitionGroup.Definitions.get_Item(sharedParameterFromFOP.Name);
                        ExternalDefinition externalDefinition = definition as ExternalDefinition;
                        FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.ParameterAndFamily.ParameterName + "_newsp").FirstOrDefault();
                        doc.FamilyManager.ReplaceParameter(fp, externalDefinition, fp.Definition.ParameterGroup, fp.IsInstance);
                        str1 = "Добавление общего параметра с правильным guid";
                        pName = definition.Name;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                    catch (Exception ex)
                    {
                        // здесь параметры, которые содержатся в семействах семейств нужно что то сделать с этим
                        //MessageBox.Show(ex.ToString());
                        //FamilyParameter fp = doc.FamilyManager.GetParameters().Where(x => x.Definition.Name == wp.ParameterAndFamily.ParameterName + "_newsp").FirstOrDefault();
                        //doc.FamilyManager.RenameParameter(fp, wp.ParameterAndFamily.ParameterName);
                        success = false;
                        str1 = "Ошибка при замене на общий параметр";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                        //System.Windows.Forms.MessageBox.Show(wp.ParameterAndFamily.ParameterName + "\n" + ex.ToString());
                    };
                    t5.Commit();
                }
                if (!success)
                {
                    using (Transaction t4 = new Transaction(doc, "Rename Parameters On Old Name"))
                    {
                        t4.Start();
                        List<FamilyParameter> fpset = doc.FamilyManager.GetParameters().ToList();
                        foreach (var fp in fpset)
                        {
                            try
                            {
                                if (fp.Definition.Name == wp.ParameterAndFamily.ParameterName + "_newsp")
                                {
                                    doc.FamilyManager.RenameParameter(fp, wp.ParameterAndFamily.ParameterName);
                                }
                                str1 = "Ошибка при замене на общий параметр";
                                pName = wp.ParameterAndFamily.ParameterName;
                                str2 = "";
                                outputReport += $"\n{str1} {pName} {str2}";
                            }
                            catch
                            {
                                success = false;
                                str1 = "Не удалось вернуть имя параметру";
                                pName = wp.ParameterAndFamily.ParameterName;
                                str2 = "";
                                outputReport += $"\n{str1} {pName} {str2}";
                            }
                        }
                        t4.Commit();
                    }
                }
            }
            return outputReport;
        }
        private string RenameFamilyParameter(Document doc, WrongParameter wp, out bool success, out string str1, out string pName, out string str2)
        {
            success = true;
            str1 = "";
            str2 = "";
            pName = "";
            string outputReport = "";
            using (Transaction t4 = new Transaction(doc, "Rename Parameters On Old Name"))
            {
                t4.Start();
                List<FamilyParameter> fpset = doc.FamilyManager.GetParameters().ToList();
                foreach (var fp in fpset)
                {
                    try
                    {
                        if (fp.Definition.Name == wp.ParameterAndFamily.ParameterName + "_newsp")
                        {
                            doc.FamilyManager.RenameParameter(fp, wp.ParameterAndFamily.ParameterName);
                        }
                        str1 = "Возврат к прежнему имени";
                        pName = fp.Definition.Name;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                    catch
                    {
                        success = false;
                        str1 = "Не вернуть имя параметру";
                        pName = wp.ParameterAndFamily.ParameterName;
                        str2 = "";
                        outputReport += $"\n{str1} {pName} {str2}";
                    }
                }
                t4.Commit();
            }
            return outputReport;
        }
        public string DeleteWP(Document doc, WrongParameter wp, out bool success, out string str1, out string pName, out string str2)
        {
            str1 = "";
            str2 = "";
            pName = "";
            success = true;
            string outputReport = "";
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
                            str1 = "Удаление общего параметра";
                            pName = wp.ParameterAndFamily.ParameterName;
                            str2 = ": " + result;
                            outputReport += $"\n{str1} {pName} {str2}";
                        }
                        catch (Exception ex)
                        {
                            success = false;
                            str1 = "Не удалось удалить общий параметр";
                            pName = wp.ParameterAndFamily.ParameterName;
                            str2 = "";
                            outputReport += $"\n{str1} {pName} {str2}";
                        }
                    }
                }
                t3.Commit();
            }
            return outputReport;
        }
        public string DeleteSP(SharedParameterElement sp, Document doc)
        {
            string output = "параметр не обнаружен и";
            if (sp != null)
            {
                try
                {
                    output = "параметр ";// + sp.GetDefinition().Name;
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
                    return output + " не удален";
                }
            }
            return output;
        }
    }
}
