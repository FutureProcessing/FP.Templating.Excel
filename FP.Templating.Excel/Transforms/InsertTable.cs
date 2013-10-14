namespace FP.Templating.Excel.Transforms
{
    using System.Linq;
    using System.Collections;

    using ClosedXML.Excel;

    public class InsertTable : ICellTransformer
    {
        public bool Recognize(IXLCell cell)
        {
            return cell.GetString().Contains('%');
        }

        public void Transform(IXLCell cell, TemplateState state)
        {
            var val = cell.Value.ToString().Substring(1);
            var collectionName = val.Substring(0, val.IndexOf('%'));

            var collection = ((IEnumerable)state.Model.GetType().GetProperty(collectionName).GetValue(state.Model, null)).OfType<object>().ToList();

            var fields = state.RemainingCells.DequeueWhile(x => x.Value.ToString().StartsWith("%" + collectionName + "%")).Select(x => new { Cell = x, Expr = x.GetString().Substring(("%" + collectionName + "%").Length + 1) }).ToList();
            if (collection.Any())
            {
                IXLRange range;
                
                if (collection.Count > 1)
                {
                    IXLRows newRows = cell.WorksheetRow().InsertRowsBelow(collection.Count - 1);
                    range = cell.Worksheet.Range(fields[0].Cell.Address, newRows.Last().LastCell().Address);
                }
                else
                {
                    range = fields[0].Cell.Worksheet.Range(fields[0].Cell.Address, fields.Last().Cell.Address);
                }                                

                for (int rowIndex = 0; rowIndex < collection.Count; rowIndex++)
                {
                    var targetRow = range.Row(rowIndex + 1);
                    for (int cellIndex = 0; cellIndex < fields.Count; cellIndex++)
                    {
                        var targetCell = targetRow.Cell(cellIndex + 1);
                        targetCell.Value = fields[cellIndex].Expr;
                        targetCell.Style = fields[cellIndex].Cell.Style;

                        state.TransformCell(targetCell, collection[rowIndex]);
                    }
                }
            }
            else
            {
                fields[0].Cell.Worksheet.Range(fields[0].Cell.Address, fields.Last().Cell.Address).Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
        }
    }
}
