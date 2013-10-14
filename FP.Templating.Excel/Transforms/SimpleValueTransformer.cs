namespace FP.Templating.Excel.Transforms
{
    using ClosedXML.Excel;

    public class SimpleValueTransformer : ICellTransformer
    {
        public bool Recognize(IXLCell cell)
        {
            return cell.Value.ToString().Contains("{");
        }

        public void Transform(IXLCell cell, TemplateState state)
        {
            cell.Value = SimpleValue(cell.Value.ToString(), state.Model);
            state.RemainingCells.Dequeue();
        }

        private static string SimpleValue(string format, object model)
        {
            return format.FormatWith(model);
        }
    }
}
