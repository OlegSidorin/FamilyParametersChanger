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
using System.Windows;

namespace FamilyParametersChanger
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class MainCommandQ : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            if (!doc.IsFamilyDocument) return Result.Succeeded;

            FilePathViewModel filePathViewModel = new FilePathViewModel();
            filePathViewModel.FileType = FileType.Family;
            filePathViewModel.PathString = doc.PathName;
            filePathViewModel.Name = doc.Title + ".rfa";

            FlowDocumentForReport flowDocumentForReport = new FlowDocumentForReport();
            flowDocumentForReport.FlowDocument = new System.Windows.Documents.FlowDocument();
            flowDocumentForReport.MakeHead(filePathViewModel);

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

            if (listOfParametersAndComments.Count != 0)
                flowDocumentForReport.MakeReportAboutWrongParameters(listOfParametersAndComments);
            else
                flowDocumentForReport.MakeReportThatAllOk();

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

            if (listOfParameterAndFamilies.Count != 0)
                flowDocumentForReport.MakeReportAboutWrongParametersWithFamilies(listOfParameterAndFamilies);
            if (listOfFamilyAndParameters.Count != 0)
                flowDocumentForReport.MakeReportAboutFamiliesAndWrongParameters(listOfFamilyAndParameters);

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

            Main.ParametersFIXList = listOfParameterAndFamilies;
            Main.FamiliesFIXList = listOfFamilyAndParameters;
            Main.ParametersWithOutFamiliesFIXList = listOfParametersWithOutFamilies;

            ReportWindow reportWindow = new ReportWindow();
            reportWindow.FileName = doc.Title;
            reportWindow.flowDocScrollViewer.Document = flowDocumentForReport.FlowDocument;
            reportWindow.Show();
            return Result.Succeeded;
        }
    }
}
