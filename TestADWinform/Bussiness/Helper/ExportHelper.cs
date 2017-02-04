using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.VariantTypes;

namespace Bussiness
{
    public class ExportHelper
    {
        protected const string Sheet1 = "Sheet1";
        private const string XNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        const int STARTROW = 1;
        const int STARTCOLUMN = 1;

        /// <summary>
        /// 创建Excel文件以及Sheet
        /// 调用示例：ExcelIO.ExcelHelper.CreateExcel(@"D:\Test2.xlsx", new string[] { "MySheet1" });
        /// </summary>
        /// <param name="fileFullName">文件全名，如：C:\TestCreateExcel.xlsx。也可以不输入，由系统自动产生文件名，并保存在系统目录下</param>
        /// <param name="sheetNames">Sheet名称</param>
        /// <param name="simpleDocument">是否创建简单的Excel文档</param>
        /// <returns>文件路径名</returns>
        public static string CreateExcel(string fileFullName, string[] sheetNames, bool simpleDocument = true)
        {
            Utilities.DealWithExistedFileName(ref fileFullName);
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullName, SpreadsheetDocumentType.Workbook, true))
            {
                CreateParts(package, sheetNames);
            }

            return fileFullName;
        }
        /// <summary>
        /// 创建一个 Excel 工作簿
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sheetNames"></param>
        protected static void CreateParts(SpreadsheetDocument document, string[] sheetNames = null)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>();
            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart, sheetNames);

            //SharedStringTablePart shareStringTablePart = CreateSharedStringTablePart(workbookPart); //不需要，在插入字符串的时候会新增 SharedStringTablePart

            //WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            //GenerateWorksheetPartContent(worksheetPart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            GenerateWorkbookStylesPartContent(workbookStylesPart);

            ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
            GenerateThemePartContent(themePart);

            SetPackageProperties(document);
        }
        private static void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Properties properties1 = new Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Application application1 = new Application();
            application1.Text = "Microsoft Excel";
            DocumentSecurity documentSecurity1 = new DocumentSecurity();
            documentSecurity1.Text = "0";
            ScaleCrop scaleCrop1 = new ScaleCrop();
            scaleCrop1.Text = "false";

            HeadingPairs headingPairs1 = new HeadingPairs();

            VTVector vTVector1 = new VTVector() { BaseType = VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Variant variant1 = new Variant();
            VTLPSTR vTLPSTR1 = new VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Variant variant2 = new Variant();
            VTInt32 vTInt321 = new VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            TitlesOfParts titlesOfParts1 = new TitlesOfParts();

            VTVector vTVector2 = new VTVector() { BaseType = VectorBaseValues.Lpstr, Size = (UInt32Value)3U };
            VTLPSTR vTLPSTR2 = new VTLPSTR();
            vTLPSTR2.Text = "Sheet1";
            VTLPSTR vTLPSTR3 = new VTLPSTR();
            vTLPSTR3.Text = "Sheet2";
            VTLPSTR vTLPSTR4 = new VTLPSTR();
            vTLPSTR4.Text = "Sheet3";

            vTVector2.Append(vTLPSTR2);
            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);

            titlesOfParts1.Append(vTVector2);
            Company company1 = new Company();
            company1.Text = "";
            LinksUpToDate linksUpToDate1 = new LinksUpToDate();
            linksUpToDate1.Text = "false";
            SharedDocument sharedDocument1 = new SharedDocument();
            sharedDocument1.Text = "false";
            HyperlinksChanged hyperlinksChanged1 = new HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            ApplicationVersion applicationVersion1 = new ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart.Properties = properties1;

        }
        /// <summary>
        /// 生成 WorkbookPart 内容
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetNames"></param>
        private static void GenerateWorkbookPartContent(WorkbookPart workbookPart, string[] sheetNames = null)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "Microsoft Office Excel" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { FilterPrivacy = true, DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 240, YWindow = 105, WindowWidth = (UInt32Value)14805U, WindowHeight = (UInt32Value)8010U };

            bookViews1.Append(workbookView1);

            Sheets sheets = new Sheets();
            if (sheetNames == null)
            {
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                GenerateWorksheetPartContent(worksheetPart);

                Sheet sheet = new Sheet()
                {
                    Name = Sheet1,
                    SheetId = 1,
                    Id = workbookPart.GetIdOfPart(worksheetPart)
                };

                sheets.Append(sheet);
            }
            else
            {
                for (int i = 0; i < sheetNames.Length; i++)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    GenerateWorksheetPartContent(worksheetPart);

                    Sheet sheet = new Sheet()
                    {
                        Name = sheetNames[i],
                        SheetId = sheets.Elements<Sheet>().Count() > 0 ? (sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1) : 1,
                        Id = workbookPart.GetIdOfPart(worksheetPart)
                    };

                    sheets.Append(sheet);
                }
            }

            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)122211U, FullCalculationOnLoad = true };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets);
            workbook1.Append(calculationProperties1);

            workbookPart.Workbook = workbook1;
            workbookPart.Workbook.Save();
        }
        /// <summary>
        /// 生成 WorksheetPart 内容
        /// </summary>
        /// <param name="worksheetPart"></param>
        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { WorkbookViewId = (UInt32Value)0U };

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D };

            SheetData sheetData1 = new SheetData();
            //Row row = new Row { RowIndex = 1 };
            //Cell cell = new Cell() { CellReference = new StringValue("A1"), CellValue = new CellValue("1"), DataType = new EnumValue<CellValues>(CellValues.Number) };
            //row.Append(cell);
            //sheetData1.Append(row);

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

            worksheet.Append(sheetDimension1);
            worksheet.Append(sheetViews1);
            worksheet.Append(sheetFormatProperties1);
            worksheet.Append(sheetData1);
            worksheet.Append(pageMargins1);

            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();
        }
        /// <summary>
        /// 生成 WorkbookStylesPart 内容
        /// </summary>
        /// <param name="workbookStylesPart"></param>
        private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet1 = new CustomStylesheet();
            workbookStylesPart.Stylesheet = stylesheet1;
        }
        /// <summary>
        /// 生成 ThemePart 内容
        /// </summary>
        /// <param name="themePart"></param>
        private static void GenerateThemePartContent(ThemePart themePart)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme2.Append(majorFont1);
            fontScheme2.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme2);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart.Theme = theme1;

        }
        /// <summary>
        /// 设置 Excel 工作簿属性
        /// </summary>
        /// <param name="document"></param>
        private static void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "段建祥";
            //document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2006-09-16T00:00:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            //document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2006-09-16T00:00:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        /// <summary>
        /// 将 实体对象列表 中的数据写入 Excel（对象方式）
        /// </summary>
        /// <typeparam name="T">实体对象类型</typeparam>
        /// <param name="fileFullName"></param>
        /// <param name="objects"></param>
        /// <param name="sheetName"></param>
        /// <param name="template"></param>
        public static void WriteExcel<T>(string fileFullName, List<T> objects, string sheetName, bool template = false)
        {
            if (File.Exists(fileFullName))
            {
                if (!template)
                    MakeCustomStylesheet(fileFullName);

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileFullName, true))
                {
                    ParameterTClass<T> param = new ParameterTClass<T>();
                    param.document = document;
                    param.objects = objects;
                    param.sheetName = sheetName;
                    param.template = template;

                    DataIntoSheetData<T>(param);
                }
            }
        }

        /// <summary>
        /// 将数据写入 DataSheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="param"></param>
        protected static void DataIntoSheetData<T>(ParameterTClass<T> param)
        {
            if (param.sheetData == null)
                param.sheetData = param.document.GetFirstSheetData(param.sheetName);

            //为单元格应用列样式做准备
            param.columns = param.document.GetWorksheet(param.sheetName).Elements<Columns>().FirstOrDefault();

            Row row = null;

            //写标题行
            param.isHeaderRow = true;
            param.rowIndex = STARTROW;
            row = DataIntoRow<T>(param);
            if (row != null)
                AddRowToSheetData(param.sheetData, row);

            //写数据行
            if (param.objects != null)
            {
                param.isHeaderRow = false;
                for (int i = 0; i < param.objects.Count; i++)
                {
                    param.rowIndex = STARTROW + 1 + i; //STARTROW是标题行，加1是数据行的开始，i是当前数据行（从0开始）
                    param.obj = param.objects[i];
                    row = DataIntoRow<T>(param);
                    if (row != null)
                        AddRowToSheetData(param.sheetData, row);
                }
            }
        }

        private static Row DataIntoRow<T>(ParameterTClass<T> param)
        {
            if (param.rowIndex < 1)
                return null;

            //设置行样式
            param.styleIndex = null;

            //获得新行或者已存在的行
            Row row = param.sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == param.rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = (uint)param.rowIndex };
                param.row = row;

                //应用行样式（如果行样式和列样式同时存在，单元格什么时候应用行样式，什么时候应用列样式）
                if (param.styleIndex != null)
                    param.row.StyleIndex = param.styleIndex;
            }
            else
            {
                param.row = row;

                //应用行样式
                if (param.styleIndex != null)
                {
                    row.StyleIndex = param.styleIndex;
                }
                else
                {
                    if (row.StyleIndex != null)
                        param.styleIndex = row.StyleIndex;
                }
            }

            //在行内新增或编辑单元格
            CellIntoRow<T>(param);


            //if (row == null)
            //{
            //    //新增行
            //    param.row = new Row() { RowIndex = (uint)param.rowIndex };

            //    //设置行样式
            //    param.styleIndex = null;
            //    //行样式（如果行样式和列样式同时存在，单元格什么时候应用行样式，什么时候应用列样式）
            //    if (param.styleIndex != null)
            //        param.row.StyleIndex = param.styleIndex;

            //    row = CreateRow<T>(param);
            //}
            //else
            //{
            //    //更新行
            //    param.row = row;

            //    //设置行样式
            //    if (row.StyleIndex != null)
            //        param.styleIndex = row.StyleIndex;

            //    UpdateRow<T>(param);
            //}

            param.styleIndex = null; //恢复样式

            return row;
        }

        private static void CellIntoRow<T>(ParameterTClass<T> param)
        {
            if (param.row == null)
                return;

            if (param.fields == null)
                param.fields = GetPropertyInfo<T>();

            Cell cell = null;
            for (int i = 0; i < param.fields.Count; i++)
            {
                param.columnIndex = STARTCOLUMN + i; //STARTCOLUMN是第一列，i是当前列（从0开始）
                param.cellIndex = Utilities.ExcelColumnIndexToName(param.columnIndex) + param.rowIndex;
                //列样式（如果有行样式，先应用行样式）
                if (param.styleIndex == null)
                    param.styleIndex = GetColumnStyle(param.columns, (uint)param.columnIndex);

                string fieldName = param.fields[i];
                if (param.isHeaderRow)
                {
                    param.content = fieldName;
                    if (!param.template)
                        //自定义 Header 样式
                        param.styleIndex = null; //12; //13
                }
                else
                {
                    if (param.obj == null)
                        break;
                    PropertyInfo pInfo = param.obj.GetType().GetProperty(fieldName);
                    param.content = pInfo.GetValue(param.obj, null);
                    //自定义样式
                    if (!param.template && param.fields[i].ToLower() == "percent")
                        param.styleIndex = null; //12; //13
                }

                cell = DataIntoCell(param);

                //（自定义）处

                //将新建的单元格插入到行
                if (cell != null)
                    AddCellToRow(param.row, cell);

                param.styleIndex = param.row.StyleIndex; //恢复到行样式
            }
        }

        private static Cell DataIntoCell<T>(ParameterTClass<T> param)
        {
            if (string.IsNullOrEmpty(param.cellIndex))
                return null;

            Cell cell = param.row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == param.cellIndex);
            if (cell == null)
            {
                //新增单元格
                cell = CreateCell<T>(param);
            }
            else
            {
                //更新单元格
                param.cell = cell;
                UpdateCell<T>(param);
            }

            return cell;
        }

        private static Cell CreateCell<T>(ParameterTClass<T> param)
        {
            if (string.IsNullOrEmpty(param.cellIndex))
                return null;

            SharedStringTablePart sharedStringTablePart = param.document.GetSharedStringTable();
            Cell cell = CreateSetCell(param.cellIndex, param.content, sharedStringTablePart, param.styleIndex, param.template);

            return cell;
        }

        private static void UpdateCell<T>(ParameterTClass<T> param)
        {
            if (param.cell == null)
                return;

            SharedStringTablePart sharedStringTablePart = param.document.GetSharedStringTable();
            UpdateSetCell(param.cell, param.content, sharedStringTablePart, param.styleIndex, param.template);
        }

        #region Object 共用的方法（DataTable, List<T>）

        /// <summary>
        /// 创建并设置单元格
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <param name="content"></param>
        /// <param name="sharedStringTablePart"></param>
        /// <param name="styleIndex"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        private static Cell CreateSetCell(string cellIndex, object content, SharedStringTablePart sharedStringTablePart, UInt32Value styleIndex, bool template)
        {
            Cell cell = null;

            if (content != null)
            {
                Type dataType = content.GetType();

                int sharedStringIndex;
                switch (dataType.Name)
                {
                    case "Boolean":
                        string val = (bool)content ? "Ture" : "False";
                        sharedStringIndex = InsertSharedStringItem(val, sharedStringTablePart);
                        cell = new TextSharedStringCell(cellIndex, sharedStringIndex.ToString(), template ? styleIndex : (styleIndex??0)); //1
                        break;
                    case "DateTime":
                        cell = new DateCell(cellIndex, (DateTime)content, template ? styleIndex : (styleIndex??2)); //3
                        break;
                    case "Int16":
                    case "Int32":
                    case "Int64":
                        cell = new NumberIntergerCell(cellIndex, content.ToString(), template ? styleIndex : (styleIndex??0)); //1
                        break;
                    case "Double":
                    case "Float":
                    case "Single":
                        cell = new NumberSingleCell(cellIndex, content.ToString(), template ? styleIndex : (styleIndex??6)); //7
                        break;
                    case "Decimal":
                        cell = new NumberDecimalCell(cellIndex, content.ToString(), template ? styleIndex : (styleIndex??8)); //9
                        break;
                    case "String":
                        sharedStringIndex = InsertSharedStringItem(content.ToString(), sharedStringTablePart);
                        cell = new TextSharedStringCell(cellIndex, sharedStringIndex.ToString(), template ? styleIndex : (styleIndex??0)); //1
                        break;
                    default:
                        cell = new TextInlineStringCell(cellIndex, content.ToString(), template ? styleIndex : (styleIndex??0)); //1
                        break;
                }
            }

            return cell;
        }

        /// <summary>
        /// 更新并设置单元格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="content"></param>
        /// <param name="sharedStringTablePart"></param>
        /// <param name="styleIndex"></param>
        /// <param name="template"></param>
        private static void UpdateSetCell(Cell cell, object content, SharedStringTablePart sharedStringTablePart, UInt32Value styleIndex, bool template)
        {
            if (content != null)
            {
                Type dataType = content.GetType();

                int sharedStringIndex;

                switch (dataType.Name)
                {
                    case "Boolean":
                        string val = (bool)content ? "Ture" : "False";
                        sharedStringIndex = InsertSharedStringItem(val, sharedStringTablePart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                        cell.CellValue = new CellValue(sharedStringIndex.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 0; //1
                        }
                        break;
                    case "DateTime":
                        cell.CellValue = new CellValue() { Text = ((DateTime)content).ToOADate().ToString() };
                        if (!template)
                        {
                            cell.StyleIndex = 2; //3
                        }
                        break;
                    case "Int16":
                    case "Int32":
                    case "Int64":
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(content.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 0; //1
                        }
                        break;
                    case "Double":
                    case "Float":
                    case "Single":
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(content.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 6; //7
                        }
                        break;
                    case "Decimal":
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(content.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 8; //9
                        }
                        break;
                    case "String":
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        sharedStringIndex = InsertSharedStringItem(content.ToString(), sharedStringTablePart);
                        cell.CellValue = new CellValue(sharedStringIndex.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 0; //1
                        }
                        break;
                    default:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        cell.CellValue = new CellValue(content.ToString());
                        if (!template)
                        {
                            cell.StyleIndex = 0; //1
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// 插入新行
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="newRow"></param>
        private static void AddRowToSheetData(SheetData sheetData, Row newRow)
        {
            if (newRow != null)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == newRow.RowIndex.Value);
                if (row != null)
                    return;

                InsertRowToSheet(sheetData, newRow);
            }
        }

        /// <summary>
        /// 插入新单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="newCell"></param>
        private static void AddCellToRow(Row row, Cell newCell)
        {
            if (newCell != null)
            {
                Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value.Length == newCell.CellReference.Value.Length && string.Compare(c.CellReference.Value, newCell.CellReference.Value, true) == 0);
                if (cell != null)
                    return;

                InsertCellToRow(row, newCell);
            }
        }

        /// <summary>
        /// 获取列样式索引值
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        private static UInt32Value GetColumnStyle(Columns columns, uint columnIndex)
        {
            UInt32Value styleIndex = null;

            if (columns != null)
            {
                Column column = columns.Elements<Column>().FirstOrDefault(c => c.Max.Value >= columnIndex && c.Min.Value <= columnIndex);
                if (column != null && column.Style != null)
                {
                    styleIndex = column.Style;
                }
            }

            return styleIndex;
        }

        /// <summary>
        /// 字符串类型的单元格，将文本插入到 SharedStringTablePart 中
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="shareStringPart"></param>
        /// <returns></returns>
        protected static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
                shareStringPart.SharedStringTable.Count = 1;
                shareStringPart.SharedStringTable.UniqueCount = 1;
            }
            int i = 0;
            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        /// 自定义列宽
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="dataTable"></param>
        protected static void SetColumnWidth(SpreadsheetDocument document, string sheetName, System.Data.DataTable dataTable)
        {
            Worksheet worksheet = document.GetWorksheet(sheetName);
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            Columns columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                sheetData.InsertBeforeSelf<Columns>(columns); //Columns元素在SheetData元素之前
            }

            Column column = null;
            uint columnIndex = 0;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                double columnWidth = 10; //dataTable.Columns[i].ColumnName.Length + 5;
                if (dataTable.Columns[i].ColumnName.ToUpper() == "GUID")
                    columnWidth = 40;

                if (dataTable.Columns[i].ColumnName.ToLower() == "birthday")
                    columnWidth = 20;

                if (dataTable.Columns[i].ColumnName.ToLower() == "birthplace")
                    columnWidth = 25;

                columnIndex = (uint)(i + 1);
                column = columns.Elements<Column>().FirstOrDefault(c => columnIndex >= c.Min.Value && columnIndex <= c.Max.Value);
                if (column == null)
                {
                    Column newColumn = new CustomColumn(columnIndex, columnIndex, columnWidth);
                    column = columns.Elements<Column>().FirstOrDefault(c => c.Min.Value > columnIndex);
                    if (column == null)
                        columns.Append(newColumn);
                    else
                        column.InsertBeforeSelf<Column>(newColumn);
                }
                else
                {
                    if (columnIndex == column.Min.Value)
                    {
                        Column newColumn = (Column)column.Clone();
                        newColumn.Max.Value = columnIndex;
                        column.Min.Value++;
                        column.InsertBeforeSelf<Column>(newColumn);
                    }
                    else if (columnIndex == column.Max.Value)
                    {
                        Column newColumn = (Column)column.Clone();
                        newColumn.Min.Value = columnIndex;
                        column.Max.Value--;
                        column.InsertAfterSelf<Column>(newColumn);
                    }
                    else {
                        Column newColumn = (Column)column.Clone();
                        newColumn.Max.Value = columnIndex - 1;
                        column.Min.Value = columnIndex;
                        column.InsertBeforeSelf<Column>(newColumn);

                        newColumn = (Column)column.Clone();
                        newColumn.Min.Value = columnIndex + 1;
                        column.Max.Value = columnIndex;
                        column.InsertAfterSelf<Column>(newColumn);
                    }
                }

                //column = columns.Elements<Column>().FirstOrDefault(c => c.Max.Value == max && c.Min.Value == min);
                //if (column == null)
                //{
                //    column = new CustomColumn(min, max, columnWidth);
                //    Column refColumn = columns.Elements<Column>().FirstOrDefault(c => c.Min.Value > max);
                //    if (refColumn == null)
                //        columns.Append(column);
                //    else
                //        columns.InsertBefore<Column>(column, refColumn);
                //}
                //else
                //{
                //    column.Width.Value = columnWidth;
                //}
            }
        }

        protected static void MakeCustomStylesheet(string fileFullName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileFullName, true))
            {
                WorkbookStylesPart workbookStylePart = document.WorkbookPart.WorkbookStylesPart;
                workbookStylePart.Stylesheet = new CustomStylesheet();
            }
        }

        private static List<string> GetPropertyInfo<T>()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            // write property names
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }

        /// <summary>
        /// 向Sheet中插入Row（OpenXml sdk Object）
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="newRow"></param>
        private static void InsertRowToSheet(SheetData sheetData, Row newRow)
        {
            Row refRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value > newRow.RowIndex.Value);
            if (refRow == null)
                sheetData.Append(newRow);
            else
                sheetData.InsertBefore<Row>(newRow, refRow);
        }

        /// <summary>
        /// 向Row中插入Cell（OpenXml sdk Object）
        /// </summary>
        /// <param name="row"></param>
        /// <param name="newCell"></param>
        private static void InsertCellToRow(Row row, Cell newCell)
        {
            Cell refCell = row.Elements<Cell>()
                                .FirstOrDefault(c => c.CellReference.Value.Length == newCell.CellReference.Value.Length && string.Compare(c.CellReference.Value, newCell.CellReference.Value, true) > 0 || c.CellReference.Value.Length > newCell.CellReference.Value.Length);
            if (refCell == null)
                row.Append(newCell);
            else
                row.InsertBefore<Cell>(newCell, refCell);
        }

        #endregion

    }

    public class ParameterTClass<T>
    {
        public DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document { get; set; }
        public DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData { get; set; }
        public DocumentFormat.OpenXml.Packaging.SharedStringTablePart sharedStringTablePart { get; set; }
        public DocumentFormat.OpenXml.Spreadsheet.Columns columns { get; set; }
        public DocumentFormat.OpenXml.Spreadsheet.Row row { get; set; }
        public DocumentFormat.OpenXml.Spreadsheet.Cell cell { get; set; }
        public DocumentFormat.OpenXml.UInt32Value styleIndex { get; set; }

        public List<T> objects { get; set; }
        public T obj { get; set; }
        public List<string> fields { get; set; }

        public bool template { get; set; }
        public string sheetName { get; set; }
        public int rowIndex { get; set; }
        public int columnIndex { get; set; }
        public string cellIndex { get; set; }
        public object content { get; set; }
        public bool isHeaderRow { get; set; }
    }

    public static class OpenXmlHelper
    {
        /// <summary>
        /// 获取 WorksheetPart
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static WorksheetPart GetWorksheetPart(this SpreadsheetDocument document, string sheetName = null)
        {
            var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
            var sheet = (sheetName == null
                             ? sheets.FirstOrDefault()
                             : sheets.FirstOrDefault(s => s.Name == sheetName)) ?? sheets.FirstOrDefault();

            return (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
        }

        /// <summary>
        /// 获取 Worksheet
        /// </summary>
        /// <param name="document">document对象</param>
        /// <param name="sheetName">sheetName可空</param>
        /// <returns>Worksheet对象</returns>
        public static Worksheet GetWorksheet(this SpreadsheetDocument document, string sheetName = null)
        {
            return document.GetWorksheetPart(sheetName).Worksheet;
        }

        /// <summary>
        /// 获取第一个SheetData
        /// </summary>
        /// <param name="document">SpreadsheetDocument对象</param>
        /// <param name="sheetName">sheetName可为空</param>
        /// <returns>SheetData对象</returns>
        public static SheetData GetFirstSheetData(this SpreadsheetDocument document, string sheetName = null)
        {
            return document.GetWorksheet(sheetName).GetFirstChild<SheetData>();
        }

        /// <summary>
        /// 获取第一个SheetData
        /// </summary>
        /// <param name="worksheet">Worksheet对象</param>
        /// <returns>SheetData对象</returns>
        public static SheetData GetFirstSheetData(this Worksheet worksheet)
        {
            return worksheet.GetFirstChild<SheetData>();
        }

        /// <summary>
        /// 获了共享字符的表格对象
        /// </summary>
        /// <param name="document">SpreadsheetDocument</param>
        /// <returns>SharedStringTablePart对角</returns>
        public static SharedStringTablePart GetSharedStringTable(this SpreadsheetDocument document)
        {
            var sharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sharedStringTablePart == null)
            {
                sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }
            return sharedStringTablePart;
        }

    }

    internal class TextSharedStringCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为字符串类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="sharedStringIndex">文本在 SharedString 中的索引，如：3</param>
        /// <param name="styleIndex"></param>
        public TextSharedStringCell(string cellIndex, string sharedStringIndex, UInt32Value styleIndex)
        {
            this.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            this.CellValue = new CellValue(sharedStringIndex);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class TextInlineStringCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为字符串类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">字符串文本</param>
        /// <param name="styleIndex"></param>
        public TextInlineStringCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.InlineString;
            this.InlineString = new InlineString { Text = new Text { Text = text } };
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class DateCell : Cell
    {
        /// <summary>
        /// /// <summary>
        /// 构造函数：内容为日期类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="dateTime">日期</param>
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <param name="dateTime"></param>
        /// <param name="styleIndex"></param>
        public DateCell(string cellIndex, DateTime dateTime, UInt32Value styleIndex)
        {
            //this.DataType = CellValues.Date;
            this.CellValue = new CellValue { Text = dateTime.ToOADate().ToString() };
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class NumberCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为 Number 类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">Number 数字的文本形式</param>
        /// <param name="styleIndex"></param>
        public NumberCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellValue = new CellValue(text);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class NumberIntergerCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为 Interger 类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">Number 数字的文本形式</param>
        /// <param name="styleIndex"></param>
        public NumberIntergerCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellValue = new CellValue(text);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class NumberSingleCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为 Number 类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">Number 数字的文本形式</param>
        /// <param name="styleIndex"></param>
        public NumberSingleCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellValue = new CellValue(text);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class NumberDecimalCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为 Number 类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">Number 数字的文本形式</param>
        /// <param name="styleIndex"></param>
        public NumberDecimalCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellValue = new CellValue(text);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class FomulaCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为公式的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">公式文本</param>
        /// <param name="styleIndex"></param>
        public FomulaCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellFormula = new CellFormula { CalculateCell = true, Text = text };
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class PercentCell : Cell
    {
        /// <summary>
        /// 内容样式为百分数形式的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="text">Number类型的数字字符</param>
        /// <param name="styleIndex"></param>
        public PercentCell(string cellIndex, string text, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Number;
            this.CellValue = new CellValue(text);
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class BooleanCell : Cell
    {
        /// <summary>
        /// 构造函数：内容为布尔类型的单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引，如：A1</param>
        /// <param name="dateTime">日期</param>
        /// <param name="styleIndex"></param>
        public BooleanCell(string cellIndex, string value, UInt32Value styleIndex)
        {
            this.DataType = CellValues.Boolean;
            this.CellValue = new CellValue { Text = value };
            this.CellReference = cellIndex;
            if (styleIndex != null)
                this.StyleIndex = styleIndex;
        }
    }

    internal class HeaderCell : TextSharedStringCell
    {
        public HeaderCell(string cellIndex, string sharedStringIndex, UInt32Value styleIndex) :
            base(cellIndex, sharedStringIndex, styleIndex)
        {
            this.StyleIndex = 11;
        }
    }

    internal class HeaderCell2 : TextInlineStringCell
    {
        public HeaderCell2(string cellIndex, string text, UInt32Value styleIndex) :
            base(cellIndex, text, styleIndex)
        {
            this.StyleIndex = 11;
        }
    }

    internal class NumberCell2 : NumberCell
    {
        public NumberCell2(string cellIndex, string text, UInt32Value styleIndex)
            : base(cellIndex, text, styleIndex)
        {
            this.StyleIndex = 5;
        }
    }

    internal class CustomColumn : Column
    {
        public CustomColumn(UInt32 startColumnIndex,
               UInt32 endColumnIndex, double columnWidth)
        {
            this.Min = startColumnIndex;
            this.Max = endColumnIndex;
            this.Width = columnWidth;
            this.CustomWidth = true;
            //this.BestFit = true;
        }
    }

    public class CustomStylesheet : Stylesheet
    {
        public CustomStylesheet()
        {
            var fonts = new Fonts();
            //Font Index(FontId) 0
            var font = new Font();
            var fontName = new FontName
            {
                Val = StringValue.FromString("Arial")
            };
            var fontSize = new FontSize
            {
                Val = DoubleValue.FromDouble(11)
            };
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);

            //Font Index(FontId) 1
            font = new Font();
            fontName = new FontName
            {
                Val = StringValue.FromString("Arial")
            };
            fontSize = new FontSize
            {
                Val = DoubleValue.FromDouble(16)
            };
            font.FontName = fontName;
            font.FontSize = fontSize;
            font.Bold = new Bold();
            font.Color = new Color() { Rgb = "FFFF0000" };
            fonts.Append(font);

            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);

            var fills = new Fills();
            //Fill index(FillId) 0
            var fill = new Fill();
            var patternFill = new PatternFill
            {
                PatternType = PatternValues.None
            };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            //Fill index(FillId) 1
            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Gray125
            };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            //Fill index(FillId) 2, 自定义的从2开始，默认0是None，1是Gray125
            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor()
            };
            patternFill.ForegroundColor = TranslateForeground(System.Drawing.Color.LightBlue);
            patternFill.BackgroundColor = new BackgroundColor
            {
                Rgb = patternFill.ForegroundColor.Rgb
            };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            //Fill index(FillId) 3
            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor()
            };
            patternFill.ForegroundColor = TranslateForeground(System.Drawing.Color.DodgerBlue);
            patternFill.BackgroundColor = new BackgroundColor
            {
                Rgb = patternFill.ForegroundColor.Rgb
            };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

            var borders = new Borders();
            //All Boarder Index(BorderId) 0
            var border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);

            //All Boarder Index(BorderId) 1
            border = new Border
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);

            //Top and Bottom Boarder Index(BorderId) 2
            border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

            var cellStyleFormats = new CellStyleFormats();
            //(FormatId) 0
            var cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellStyleFormats.Append(cellFormat);

            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);

            var cellFormats = new CellFormats();
            // index 0
            cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };
            cellFormats.Append(cellFormat);

            uint iExcelIndex = 164;
            var numberingFormats = new NumberingFormats();

            var nformatDateTime = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("yyyy-MM-dd HH:mm:ss")
            };
            numberingFormats.Append(nformatDateTime);

            var nformat4Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.0000")
            };
            numberingFormats.Append(nformat4Decimal);

            var nformat3Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.000")
            };
            numberingFormats.Append(nformat3Decimal);

            var nformat2Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.00")
            };
            numberingFormats.Append(nformat2Decimal);

            var nformatInteger = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0")
            };
            numberingFormats.Append(nformatInteger);

            var nformatForcedText = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
                FormatCode = StringValue.FromString("@")
            };
            numberingFormats.Append(nformatForcedText);

            var nformatPercent = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
                FormatCode = StringValue.FromString("#0.0%")
            };
            numberingFormats.Append(nformatPercent);

            // index 1
            cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // index 2
            // Cell Standard Date format（无边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = 14, //单元格标准日期格式——NumberFormatId 为14
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // index 3
            // Cell Standard Date format（有边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = 14, //单元格标准日期格式——NumberFormatId 为14
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 4
            // Cell Date time custom format（无边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatDateTime.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 5
            // Cell Date time custom format（有边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatDateTime.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 6
            // Cell Standard Number format with 2 decimal placing（无边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true) //指示是否应用该自定义的 NumberingFormat
            };
            cellFormats.Append(cellFormat);

            // Index 7
            // Cell Standard Number format with 2 decimal placing（有边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 8
            // Cell 4 decimal custom format（无边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformat3Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 9
            // Cell 4 decimal custom format（有边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformat3Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 10
            // Cell forced number text custom format（无边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 11
            // Cell forced number text custom format（有边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 12
            // Coloured cell text（字体样式、有填充）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 1,
                FillId = 3,
                BorderId = 0,
                FormatId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyFill = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);

            // Index 13
            // Coloured cell text（字体样式、有填充、边框）
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 1,
                FillId = 3,
                BorderId = 1,
                FormatId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyFill = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);


            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            this.Append(numberingFormats);
            this.Append(fonts);
            this.Append(fills);
            this.Append(borders);
            this.Append(cellStyleFormats);
            this.Append(cellFormats);

            var css = new CellStyles();

            var cs = new CellStyle
            {
                Name = StringValue.FromString("Normal"),
                FormatId = 0,
                BuiltinId = 0
            };
            css.Append(cs);

            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            this.Append(css);

            var dfs = new DifferentialFormats
            {
                Count = 0
            };
            this.Append(dfs);

            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
            this.Append(tss);

        }

        private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new ForegroundColor()
            {
                Rgb = new HexBinaryValue()
                {
                    Value =
                        System.Drawing.ColorTranslator.ToHtml(
                        System.Drawing.Color.FromArgb(
                            fillColor.A,
                            fillColor.R,
                            fillColor.G,
                            fillColor.B)).Replace("#", "")
                }
            };
        }
    }
}
