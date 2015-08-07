# OpenXmlSpreadsheet MediaTypeFormatter
Export to Excel a data from your WebAPI OData endpoint.

# Usage

Add the code to your Global.asax:

    ExportToExcelConfig.Register(GlobalConfiguration.Configuration);

And implement ExportToExcelConfig:

    public static class ExportToExcelConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // configure Export Invoice Entries to Excel for Finance Administrator
            {
                var invoicesAdminExportToExcel = new OpenXmlSpreadsheetFormatter<Invoice>("Invoices");
                
                invoicesAdminExportToExcel.AddColumn("Invoice Number", i => i.Number, width: 206 / 7.25);
                invoicesAdminExportToExcel.AddColumn("Invoice Amount", i => i.Amount.ToString() + i.Currency.Code, width: 206 / 7.25);
                // configure query string mapping, so export will be available at http://.../yourodataendpoint/...?$format=spreadsheetmladmin
                invoicesAdminExportToExcel.MediaTypeMappings.Add(new QueryStringMapping("$format", "spreadsheetmladmin", OpenXmlSpreadsheetFormatterStatic.ContentType));
                
                // insert this formatter as 1st, because json will be used instead as it doesn't care about $format=...
                config.Formatters.Insert(0, invoicesAdminExportToExcel);
            }
        }
    }
