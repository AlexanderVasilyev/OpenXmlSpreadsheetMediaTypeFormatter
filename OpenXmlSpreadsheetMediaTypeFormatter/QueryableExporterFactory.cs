using System;
using System.Linq;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public class QueryableExporterFactory<T> : IObjectExporterFactory<T>
    {
        public IObjectExporter<T> GetNewInstance()
        {
            return new QueryableExporter<T>();
        }

        public bool SupportsType(Type type)
        {
            return typeof(IQueryable<T>).IsAssignableFrom(type);
        }
    }
}