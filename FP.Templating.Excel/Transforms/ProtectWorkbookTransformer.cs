namespace FP.Templating.Excel.Transforms
{   
    public class ProtectWorkbookTransformer : TemplateTransformer
    {
        private readonly string _password;

        public ProtectWorkbookTransformer(string password)
        {
            this._password = password;
        }

        public override void EndWorksheetProcess(ClosedXML.Excel.IXLWorksheet worksheet)
        {
            worksheet.Protect(this._password);
        }
    }
}
