using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public class OpenXmlSpreadsheetFormatter<T> : MediaTypeFormatter
    {
        public string SheetName { get; set; }
        public string FileName { get; set; }
        public List<ExportProperty<T>> ExportProperties { get; set; }
        public IObjectExporterFactory<T> ObjectExporterFactory { get; set; }
        public string CreatorUserName { get; set; }

        public OpenXmlSpreadsheetFormatter(string name = null)
        {
            if (string.IsNullOrEmpty(name))
            {
                SheetName = "Sheet 0";
                FileName = "Export to Excel";
            }
            else
            {
                SheetName = name;
                FileName = name;
            }
            CreatorUserName = "System";
            ObjectExporterFactory = new QueryableExporterFactory<T>();
            ExportProperties = new List<ExportProperty<T>>();
        }

        public void AddColumn(string name, Func<T,object> dataExtractorFunc, string title = null, string format = null, ExportValueType valueType = ExportValueType.Auto, double? width = null)
        {
            ExportProperties.Add(new ExportProperty<T>
                {
                    Name = name,
                    DataExtractionFunc = dataExtractorFunc,
                    Title = title,
                    Format = format,
                    ValueType = valueType,
                    Width = width
                });
        }

        public override bool CanReadType(Type type)
        {
            return false;
        }

        public override bool CanWriteType(Type type)
        {
            return ObjectExporterFactory.SupportsType(type);
        }

        public override MediaTypeFormatter GetPerRequestFormatterInstance(Type type, HttpRequestMessage request, MediaTypeHeaderValue mediaType)
        {
            return this;
        }

        public override Task<object> ReadFromStreamAsync(Type type, System.IO.Stream readStream, HttpContent content, IFormatterLogger formatterLogger)
        {
            throw new NotSupportedException();
        }

        public override void SetDefaultContentHeaders(Type type, HttpContentHeaders headers, MediaTypeHeaderValue mediaType)
        {
            headers.ContentType = OpenXmlSpreadsheetFormatterStatic.MediaTypeHeader;
            headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") { FileName = FileName + ".xlsx" };
        }

        public static Task RunSynchronously(Action action, CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return new TaskCompletionSource<Object>().Task;
            try
            {
                action();
                TaskCompletionSource<Object> completionSource = new TaskCompletionSource<Object>();
                completionSource.SetResult(null);
                return completionSource.Task;
            }
            catch (Exception exception)
            {
                TaskCompletionSource<Object> completionSource = new TaskCompletionSource<Object>();
                completionSource.SetException(exception);
                return completionSource.Task;
            }
        }

        public override Task WriteToStreamAsync(Type type, object value, System.IO.Stream writeStream, HttpContent content, System.Net.TransportContext transportContext)
        {
            if (type == (Type)null)
                throw new ArgumentNullException("type");
            if (writeStream == null)
                throw new ArgumentNullException("writeStream");
            return RunSynchronously(() => WriteToStream(type, value, writeStream, content, transportContext), new CancellationToken());
        }

        public void WriteToStream(Type type, object value, Stream writeStream, HttpContent content, System.Net.TransportContext transportContext)
        {
            //writeStream.Write(new[] { (byte)'t', (byte)'e', (byte)'s', (byte)'t' }, 0, 4);
            //return;
            var objectExporter = ObjectExporterFactory.GetNewInstance();
            objectExporter.ObjectToExport = value;

            var memoryStream = new MemoryStream();

            var cellFormatIndexes = new Dictionary<ExportProperty<T>, UInt32Value>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                var extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1, SheetName);

                var workbookPart1 = document.AddWorkbookPart();
                GenerateWorkbookPart1Content(workbookPart1, SheetName);

                var workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId5");
                GenerateWorkbookStylesPart1Content(workbookStylesPart1, cellFormatIndexes, ExportProperties);

                var worksheetPart3 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
                GenerateWorksheetPart3Content(worksheetPart3, objectExporter, cellFormatIndexes, ExportProperties);

                var themePart1 = workbookPart1.AddNewPart<ThemePart>("rId4");
                GenerateThemePart1Content(themePart1);

                SetPackageProperties(document, CreatorUserName);
            }

            memoryStream.WriteTo(writeStream);
        }

        private static void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1, string sheetName)
        {
            var properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            var application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            var documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            var scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            var headingPairs1 = new Ap.HeadingPairs();

            var vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            var variant1 = new Vt.Variant();
            var vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            var variant2 = new Vt.Variant();
            var vTInt321 = new Vt.VTInt32 { Text = "1" };

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            var titlesOfParts1 = new Ap.TitlesOfParts();

            var vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            var vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = sheetName;

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            var company1 = new Ap.Company();
            company1.Text = "KPMG";
            var linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            var sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            var hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            var applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

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

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private static void GenerateWorkbookPart1Content(WorkbookPart workbookPart1, string sheetName)
        {
            var workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            var fileVersion1 = new FileVersion { ApplicationName = "xl", LastEdited = "4", LowestEdited = "4", BuildVersion = "4506" };
            var workbookProperties1 = new WorkbookProperties { DefaultThemeVersion = 124226U };

            var bookViews1 = new BookViews();
            var workbookView1 = new WorkbookView { XWindow = 0, YWindow = 105, WindowWidth = 19155U, WindowHeight = 12300U };

            bookViews1.Append(workbookView1);

            var sheets1 = new Sheets();
            var sheet1 = new Sheet { Name = sheetName, SheetId = 1U, Id = "rId1" };

            sheets1.Append(sheet1);
            var calculationProperties1 = new CalculationProperties { CalculationId = 125725U };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);

            workbookPart1.Workbook = workbook1;
        }

        private static void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1, Dictionary<ExportProperty<T>, UInt32Value> cellFormatIndexes, IEnumerable<ExportProperty<T>> exportProperties)
        {
            var stylesheet1 = new Stylesheet();

            var numberingFormats = new NumberingFormats { Count = 2U };

            numberingFormats.Append(new NumberingFormat { NumberFormatId = (UInt32Value)1U, FormatCode = "yyyy\\.mm\\.dd" });

            numberingFormats.Append(new NumberingFormat { NumberFormatId = (UInt32Value)2U, FormatCode = "yyyy\\.mm\\.dd hh:MM:ss" });

            stylesheet1.NumberingFormats = numberingFormats;

            var fonts1 = new Fonts { Count = 2U };

            fonts1.Append(new Font
                {
                    Color = new Color { Theme = (UInt32Value)1U },
                    FontName = new FontName { Val = "Calibri" },
                    FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
                    FontScheme = new FontScheme { Val = FontSchemeValues.Minor },
                    FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 11D }
                });

            fonts1.Append(new Font
            {
                Bold = new Bold(),
                Color = new Color() { Theme = (UInt32Value)1U },
                FontName = new FontName() { Val = "Calibri" },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
                FontScheme = new FontScheme() { Val = FontSchemeValues.Minor },
                FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 11D }
            });

            var fills1 = new Fills { Count = 2U };
            fills1.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            fills1.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });

            var borders1 = new Borders { Count = 1U };
            borders1.Append(new Border { LeftBorder = new LeftBorder(), RightBorder = new RightBorder(), TopBorder = new TopBorder(), BottomBorder = new BottomBorder(), DiagonalBorder = new DiagonalBorder() });

            var cellStyleFormats1 = new CellStyleFormats { Count = 1U };
            cellStyleFormats1.Append(new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U });

            var cellFormats = new CellFormats();
            // default text format
            cellFormats.Append(new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, Alignment = new Alignment() { WrapText = true } });
            // header row format
            cellFormats.Append(new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true });
            // default date format
            cellFormats.Append(new CellFormat() { NumberFormatId = (UInt32Value)1U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true });
            // default datetime format
            cellFormats.Append(new CellFormat() { NumberFormatId = (UInt32Value)2U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true });
            // custom formats
            {
                cellFormatIndexes.Clear();
                uint cellFormatsCounter = (uint)cellFormats.Elements().Count();
                uint numberingFormatsCounter = numberingFormats.Count;
                foreach (var property in exportProperties.Where(a => a.Visible))
                    if (!string.IsNullOrEmpty(property.Format))
                    {
                        cellFormatIndexes[property] = cellFormatsCounter;
                        cellFormatsCounter++;
                        numberingFormatsCounter++;
                        numberingFormats.Append(new NumberingFormat { NumberFormatId = numberingFormatsCounter, FormatCode = property.Format });
                        cellFormats.Append(new CellFormat() { NumberFormatId = numberingFormatsCounter, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true });
                    }
                numberingFormats.Count = numberingFormatsCounter;
                cellFormats.Count = cellFormatsCounter;
            }

            var cellStyles1 = new CellStyles { Count = 1U };
            cellStyles1.Append(new CellStyle { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U });

            var differentialFormats1 = new DifferentialFormats { Count = 0U };
            var tableStyles1 = new TableStyles { Count = 0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;
            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }
            return columnName;
        }

        protected static Boolean IsInteger(Object v)
        {
            return v is Int16 || v is Int32 || v is Int64;
        }

        protected static Boolean IsDecimal(Object v)
        {
            return v is Decimal || v is Double;
        }

        protected static Boolean IsDate(Object v)
        {
            return v is DateTime;
        }

        /// <summary>
        /// Remove illegal XML characters from a string.
        /// </summary>
        public static string SanitizeStringForXml(string s)
        {
            if (s == null)
                return null;

            var buffer = new StringBuilder(s.Length);

            foreach (char c in s)
                if (IsLegalXmlChar(c))
                    buffer.Append(c);

            return buffer.ToString();
        }

        /// <summary>
        /// Whether a given character is allowed by XML 1.0.
        /// </summary>
        public static bool IsLegalXmlChar(int character)
        {
            return
            (
                 character == 0x9 /* == '\t' == 9   */          ||
                 character == 0xA /* == '\n' == 10  */          ||
                 character == 0xD /* == '\r' == 13  */          ||
                (character >= 0x20 && character <= 0xD7FF) ||
                (character >= 0xE000 && character <= 0xFFFD) ||
                (character >= 0x10000 && character <= 0x10FFFF)
            );
        }

        private static void GenerateWorksheetPart3Content<U>(WorksheetPart worksheetPart3, IObjectExporter<U> objectExporter, Dictionary<ExportProperty<U>, UInt32Value> cellFormatIndexes, IEnumerable<ExportProperty<U>> exportProperties)
        {
            var worksheet3 = new Worksheet();
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var columns = new Columns();

            var sheetViews3 = new SheetViews();
            sheetViews3.Append(new SheetView { TabSelected = true, WorkbookViewId = (UInt32Value)0U });

            var sheetData3 = new SheetData();

            int colNumber = exportProperties.Count(a => a.Visible);
            int rowNumber = objectExporter.Count;

            // Header
            {
                var headerRow = new Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = "1:" + colNumber.ToString() } };
                uint i = 0;
                foreach (var p in exportProperties.Where(a => a.Visible))
                {
                    headerRow.Append(new Cell { CellReference = GetColumnName((int)++i) + "1", CellValue = new CellValue(p.Title), StyleIndex = (UInt32Value)1U, DataType = CellValues.String });
                    DoubleValue columnWidth = 3 + Math.Max(3, Math.Min(50, p.Title.Length));
                    if (p.Width.HasValue)
                        columnWidth = p.Width.Value;
                    columns.Append(new Column { Width = columnWidth, Min = i, Max = i, CustomWidth = true });
                }
                sheetData3.Append(headerRow);
            }

            // Items
            {
                uint j = 1;
                foreach (var rowObject in objectExporter)
                {
                    var itemRow = new Row { RowIndex = UInt32Value.FromUInt32(++j), Spans = new ListValue<StringValue>() { InnerText = "1:" + colNumber.ToString() } };
                    var i = 0;
                    foreach (var property in exportProperties.Where(a => a.Visible))
                    {
                        var cell = new Cell { CellReference = GetColumnName(++i) + j.ToString(CultureInfo.InvariantCulture) };

                        object value = property.ExtractData(rowObject);
                        switch (property.ValueType)
                        {
                            case ExportValueType.Auto:
                                if (IsInteger(value))
                                    goto case ExportValueType.Integer;
                                if (IsDecimal(value))
                                    goto case ExportValueType.Decimal;
                                if (IsDate(value))
                                    goto case ExportValueType.Date;
                                goto case ExportValueType.Text;
                            case ExportValueType.Date:
                                cell.DataType = CellValues.Number;
                                cell.StyleIndex = cellFormatIndexes.ContainsKey(property) ? cellFormatIndexes[property] : 2;
                                cell.CellValue = new CellValue(value == null ? string.Empty : ((DateTime)value).ToOADate().ToString());
                                break;
                            case ExportValueType.DateTime:
                                cell.DataType = CellValues.Number;
                                cell.StyleIndex = cellFormatIndexes.ContainsKey(property) ? cellFormatIndexes[property] : 3;
                                cell.CellValue = new CellValue(value == null ? string.Empty : ((DateTime)value).ToOADate().ToString());
                                break;
                            case ExportValueType.Integer:
                            case ExportValueType.Decimal:
                                cell.DataType = CellValues.Number;
                                cell.StyleIndex = cellFormatIndexes.ContainsKey(property) ? cellFormatIndexes[property] : 0;
                                cell.CellValue = new CellValue(value == null ? string.Empty : value.ToString());
                                break;
                            case ExportValueType.Text:
                            default:
                                cell.DataType = CellValues.String;
                                cell.StyleIndex = cellFormatIndexes.ContainsKey(property) ? cellFormatIndexes[property] : 0;
                                cell.CellValue = new CellValue(value == null ? string.Empty : SanitizeStringForXml(value.ToString()));
                                break;
                        }
                        itemRow.Append(cell);
                    }
                    sheetData3.Append(itemRow);
                }
            }

            worksheet3.Append(new SheetDimension { Reference = "A1:" + GetColumnName(colNumber) + (rowNumber + 1).ToString() });
            worksheet3.Append(sheetViews3);
            worksheet3.Append(new SheetFormatProperties { DefaultRowHeight = 15D });
            worksheet3.Append(columns);
            worksheet3.Append(sheetData3);
            worksheet3.Append(new AutoFilter { Reference = "A1:" + GetColumnName(colNumber) + (rowNumber + 1).ToString() });
            worksheet3.Append(new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D });

            worksheetPart3.Worksheet = worksheet3;
        }

        private static void GenerateThemePart1Content(ThemePart themePart1)
        {
            var theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new A.ThemeElements();

            var colorScheme1 = new A.ColorScheme() { Name = "Office" };

            var dark1Color1 = new A.Dark1Color();
            var systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            var light1Color1 = new A.Light1Color();
            var systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            var dark2Color1 = new A.Dark2Color();
            var rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

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

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);

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

            themePart1.Theme = theme1;
        }

        private static void SetPackageProperties(OpenXmlPackage document, string creatorUserName)
        {
            document.PackageProperties.Creator = creatorUserName;
            document.PackageProperties.Created = DateTime.Now;
            document.PackageProperties.Modified = DateTime.Now;
            document.PackageProperties.LastModifiedBy = creatorUserName;
        }
    }
}