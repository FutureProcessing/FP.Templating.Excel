namespace FP.Templating.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.IO;

    using ClosedXML.Excel;

    using FP.Templating.Excel.Transforms;

    public class ExcelTemplateRunner
    {
        public List<TemplateTransformer> Pipeline { get; private set; }

        public List<ICellTransformer> CellTransformers { get; private set; }

        public ExcelTemplateRunner()
        {
            this.Pipeline = new List<TemplateTransformer>();
            this.CellTransformers = new List<ICellTransformer>(new ICellTransformer[] { new InsertTable(), new SimpleValueTransformer() });
        }

        public void Run<TModel>(TModel model, Stream templateStream, Stream outputStream)
        {
            var workbook = new XLWorkbook(templateStream);

            this.RunPipeline(t => t.BeginWorkbookProcess(workbook));
           
            foreach (var sheet in workbook.Worksheets)
            {
                this.TransformWorksheet(model, sheet); 
            }

            this.RunPipeline(t => t.EndWorkbookProcess(workbook));

            workbook.SaveAs(outputStream);
        }

        private void TransformWorksheet<TModel>(TModel model, IXLWorksheet sheet)
        {
            this.RunPipeline(t => t.BeginWorksheetProcess(sheet));

            var state = new TemplateState {Model = model};

            var inserTable = new InsertTable();
            var simpleValue = new SimpleValueTransformer();

            foreach (var row in sheet.RowsUsed().Reverse())
            {
                state.RemainingCells = new Queue<IXLCell>(row.CellsUsed());

                while (state.RemainingCells.Any())
                {
                    var cell = state.RemainingCells.Peek();

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

            this.RunPipeline(t => t.EndWorksheetProcess(sheet));
        }

        private void RunPipeline(Action<TemplateTransformer> action)
        {
            foreach (var templateTransformer in this.Pipeline)
            {
                action(templateTransformer);
            }
        }
    }
}
