using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace FamilyParametersChanger
{
    public class FlowDocumentForReport
    {
        public FlowDocument FlowDocument { get; set; }
        Thickness left_0
        {
            get
            {
                return new Thickness(0);
            }
        }
        Thickness left_7
        {
            get
            {
                return new Thickness(7, 0, 0, 0);
            }
        }
        Thickness left_25
        {
            get
            {
                return new Thickness(25,0,0,0);
            }
        }
        Thickness left_0_padding_5
        {
            get
            {
                return new Thickness(0,5,0,5);
            }
        }
        Thickness left_0_padding_25
        {
            get
            {
                return new Thickness(0, 25, 0, 25);
            }
        }
        System.Windows.Media.FontFamily fontFamilyBase
        {
            get
            {
                return new System.Windows.Media.FontFamily("Verdana");
            }
        }
        int fontSizeBase
        {
            get
            {
                return 12;
            }
        }
        int fontSizeHead
        {
            get
            {
                return 14;
            }
        }

        internal void MakeHead(FilePathViewModel filePathVM)
        {
            FlowDocument.TextAlignment = TextAlignment.Left;

            Paragraph pTitle = new Paragraph();
            pTitle.Margin = left_0;
            Run runTitle01 = new Run();
            runTitle01.Text = $"Документ: ";
            runTitle01.FontFamily = fontFamilyBase;
            runTitle01.FontWeight = FontWeights.Bold;
            runTitle01.FontSize = fontSizeBase;
            pTitle.Inlines.Add(runTitle01);
            Run runTitle02 = new Run();
            runTitle02.Text = $"{filePathVM.Name}";
            runTitle02.FontFamily = fontFamilyBase;
            runTitle02.FontWeight = FontWeights.Normal;
            runTitle02.FontSize = fontSizeBase;
            pTitle.Inlines.Add(runTitle02);
            FlowDocument.Blocks.Add(pTitle);

            Paragraph pDate = new Paragraph();
            pDate.Margin = left_0;
            Run runDate01 = new Run();
            runDate01.Text = $"Дата проверки: ";
            runDate01.FontFamily = fontFamilyBase;
            runDate01.FontWeight = FontWeights.Bold;
            runDate01.FontSize = fontSizeBase;
            pDate.Inlines.Add(runDate01);
            Run runDate02 = new Run();
            runDate02.Text = $"{DateTime.Now}";
            runDate02.FontFamily = fontFamilyBase;
            runDate02.FontWeight = FontWeights.Normal;
            runDate02.FontSize = fontSizeBase;
            pDate.Inlines.Add(runDate02);
            FlowDocument.Blocks.Add(pDate);

        }

        internal void MakeReportThatAllOk()
        {
            Paragraph pHead = new Paragraph();
            pHead.Margin = left_0_padding_25;
            Run run = new Run();
            run.Text = "Параметров не из ФОП М1 не обнаружено, все ОК";
            run.FontWeight = FontWeights.Bold;
            run.FontSize = fontSizeHead;
            run.FontFamily = fontFamilyBase;
            pHead.Inlines.Add(run);
            FlowDocument.Blocks.Add(pHead);
        }

        // для создания главного отчета
        internal void MakeReportAboutWrongParameters(List<ParameterAndFamily> listOfParametersAndComments)
        {
            Paragraph pHead = new Paragraph();
            pHead.Margin = left_0_padding_5;
            Run run = new Run();
            run.Text = "Параметры не из ФОП M1, не рекомендуется использовать:";
            run.FontWeight = FontWeights.Bold;
            run.FontSize = fontSizeHead;
            run.FontFamily = fontFamilyBase;
            pHead.Inlines.Add(run);
            FlowDocument.Blocks.Add(pHead);

            foreach (ParameterAndFamily pf in listOfParametersAndComments)
            {
                Paragraph pBody = new Paragraph();
                pBody.Margin = left_7;
                Run runName = new Run();
                runName.Text = $"{pf.ParameterName}";
                runName.FontFamily = fontFamilyBase;
                runName.FontWeight = FontWeights.Bold;
                runName.FontSize = fontSizeBase;
                pBody.Inlines.Add(runName);
                Run runReport = new Run();
                runReport.Text = $" - {pf.Comment}";
                runReport.FontFamily = fontFamilyBase;
                runReport.FontWeight = FontWeights.Normal;
                runReport.FontSize = fontSizeBase;
                pBody.Inlines.Add(runReport);
                FlowDocument.Blocks.Add(pBody);
            }
        }
        internal void MakeReportAboutWrongParametersWithFamilies(List<ParameterAndFamily> listOfParameterAndFamilies)
        {
            Paragraph pBody101 = new Paragraph();
            pBody101.Margin = left_0_padding_5;
            Run run101 = new Run
            {
                Text = "Параметры со списком семейств, в которых он был обнаружен:",
                FontFamily = fontFamilyBase,
                FontWeight = FontWeights.Bold,
                FontSize = fontSizeHead
            };
            pBody101.Inlines.Add(run101);
            FlowDocument.Blocks.Add(pBody101);

            foreach (var item in listOfParameterAndFamilies)
            {
                Paragraph pBody01 = new Paragraph();
                pBody01.Margin = left_7;
                Run run011 = new Run
                {
                    Text = $"Параметр ",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Normal,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run011);
                Run run012 = new Run()
                {
                    Text = $"{item.ParameterName} ({item.Cause.ToFriendlyString()})",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Bold,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run012);
                Run run013 = new Run
                {
                    Text = $" содержится в семействах:",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Normal,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run013);
                FlowDocument.Blocks.Add(pBody01);
                Paragraph pBody02 = new Paragraph();
                pBody02.Margin = left_7;
                Run run02 = new Run();
                run02.Text = $"{item.GetFamiliesInOneSting()}";
                run02.FontFamily = fontFamilyBase;
                run02.FontWeight = FontWeights.Normal;
                run02.FontStyle = FontStyles.Italic;
                run02.FontSize = fontSizeBase;
                pBody02.Inlines.Add(run02);
                FlowDocument.Blocks.Add(pBody02);
            }
        }
        internal void MakeReportAboutFamiliesAndWrongParameters(List<ParameterAndFamily> listOfFamilyAndParameters)
        {
            Paragraph pBody202 = new Paragraph()
            {
                Margin = left_0_padding_5,
            };
            Run run202 = new Run
            {
                Text = "Семейства со списком нежелательных параметров:",
                FontFamily = fontFamilyBase,
                FontWeight = FontWeights.Bold,
                FontSize = fontSizeHead
            };
            pBody202.Inlines.Add(run202);
            FlowDocument.Blocks.Add(pBody202);

            foreach (var item in listOfFamilyAndParameters)
            {
                Paragraph pBody01 = new Paragraph();
                pBody01.Margin = left_7;
                Run run011 = new Run
                {
                    Text = $"В семействе ",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Normal,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run011);
                Run run012 = new Run()
                {
                    Text = $"{item.FamilyName}",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Bold,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run012);
                Run run013 = new Run
                {
                    Text = $" содержатся нежелательные параметры:",
                    FontFamily = fontFamilyBase,
                    FontWeight = FontWeights.Normal,
                    FontSize = fontSizeBase
                };
                pBody01.Inlines.Add(run013);
                FlowDocument.Blocks.Add(pBody01);
                Paragraph pBody02 = new Paragraph();
                pBody02.Margin = left_25;
                Run run02 = new Run();
                run02.Text = $"{item.GetParametersInOneSting()}";
                run02.FontFamily = fontFamilyBase;
                run02.FontWeight = FontWeights.Normal;
                run02.FontStyle = FontStyles.Italic;
                run02.FontSize = fontSizeBase;
                pBody02.Inlines.Add(run02);
                FlowDocument.Blocks.Add(pBody02);
            }
        }

        internal void AddHead(string inputString)
        {
            Paragraph pHead = new Paragraph();
            pHead.Margin = left_0_padding_5;
            Run run = new Run();
            run.Text = inputString;
            run.FontWeight = FontWeights.Bold;
            run.FontSize = fontSizeHead;
            run.FontFamily = fontFamilyBase;
            pHead.Inlines.Add(run);
            FlowDocument.Blocks.Add(pHead);
        }
        internal void AddParagraph(string inputString1, string pName, string inputString2)
        {
            Paragraph pBody = new Paragraph();
            pBody.Margin = left_7;
            Run run011 = new Run
            {
                Text = inputString1 + " ",
                FontFamily = fontFamilyBase,
                FontWeight = FontWeights.Normal,
                FontSize = fontSizeBase
            };
            pBody.Inlines.Add(run011);
            Run run012 = new Run()
            {
                Text = pName,
                FontFamily = fontFamilyBase,
                FontWeight = FontWeights.Bold,
                FontSize = fontSizeBase
            };
            pBody.Inlines.Add(run012);
            Run run013 = new Run
            {
                Text = " " + inputString2,
                FontFamily = fontFamilyBase,
                FontWeight = FontWeights.Normal,
                FontSize = fontSizeBase
            };
            pBody.Inlines.Add(run013);
            FlowDocument.Blocks.Add(pBody);
        }


    }
}
