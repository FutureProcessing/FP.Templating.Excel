namespace FP.Templating.Excel
{
    using System.Collections.Generic;
    using System.Linq;

    using ClosedXML.Excel;

    public class TemplateState
    {
        public object Model { get; set; }

        public Queue<IXLCell> RemainingCells { get; set; }

        public Queue<ICellTransformer> CellTransformers { get; set; }

        public void TransformCell(IXLCell cell, object model)
        {
            var state = new TemplateState { Model = model, RemainingCells = new Queue<IXLCell>(new[] { cell }) };

            var cellPipeline = new Queue<ICellTransformer>(this.CellTransformers);

            while (cellPipeline.Any() && !cellPipeline.Peek().Recognize(cell))
            {
                cellPipeline.Dequeue();
            }

            state.CellTransformers = cellPipeline;

            if (cellPipeline.Any())
            {
                var transformer = cellPipeline.Dequeue();

                transformer.Transform(cell, state);
            }
            else
            {
                state.RemainingCells.Dequeue();
            }
        }
    }
}