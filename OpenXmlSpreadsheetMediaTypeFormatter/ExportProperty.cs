using System;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public class ExportProperty<T>
    {
        private string _title;

        public ExportProperty()
        {
            ValueType = ExportValueType.Auto;
            Visible = true;
        }

        public bool Visible { get; set; }

        public string Name { get; set; }

        public string Title { get { return _title ?? Name; } set { _title = value; } }

        public string Format { get; set; }

        public ExportValueType ValueType { get; set; }

        public double? Width { get; set; }

        public Func<T, object> DataExtractionFunc { get; set; }

        public object ExtractData(object fromObject)
        {
            return DataExtractionFunc((T) fromObject);
        }
    }
}