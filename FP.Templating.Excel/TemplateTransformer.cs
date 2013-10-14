namespace FP.Templating.Excel
{
    using ClosedXML.Excel;

    public abstract class TemplateTransformer
    {
        public virtual void BeginWorkbookProcess(XLWorkbook workbook)
        {
        }

        public virtual void BeginWorksheetProcess(IXLWorksheet worksheet)
        {
        }

        public virtual void EndWorksheetProcess(IXLWorksheet worksheet)
        {
        }

        public virtual void EndWorkbookProcess(XLWorkbook workbook)
        {
        }
    }
}