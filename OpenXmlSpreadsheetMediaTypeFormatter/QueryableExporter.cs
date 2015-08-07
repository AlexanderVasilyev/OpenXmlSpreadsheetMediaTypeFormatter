using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public class QueryableExporter<T> : IObjectExporter<T>
    {
        private List<T> _items;
        public IEnumerator<T> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_items).GetEnumerator();
        }

        public object ObjectToExport { set { _items = ((IQueryable<T>)value).ToList(); } }
        public int Count { get { return _items.Count; } }
    }
}