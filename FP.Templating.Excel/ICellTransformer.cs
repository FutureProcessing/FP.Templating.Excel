namespace FP.Templating.Excel
{
    using ClosedXML.Excel;

    public interface ICellTransformer
    {
        bool Recognize(IXLCell cell);

        void Transform(IXLCell cell, TemplateState state);
    }
}