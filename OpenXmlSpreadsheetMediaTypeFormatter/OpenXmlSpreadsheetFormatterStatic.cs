using System.Net.Http.Headers;

namespace OpenXmlSpreadsheetMediaTypeFormatter
{
    public static class OpenXmlSpreadsheetFormatterStatic
    {
        public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public static readonly MediaTypeHeaderValue MediaTypeHeader = new MediaTypeHeaderValue(ContentType);
    }
}