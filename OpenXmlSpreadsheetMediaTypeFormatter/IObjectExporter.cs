using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public interface IObjectExporter<T> : IEnumerable<T>
    {
        object ObjectToExport { set; }
        int Count { get; }
    }

    public interface IObjectExporterFactory<T>
    {
        IObjectExporter<T> GetNewInstance();
        bool SupportsType(Type type);
    }
}
