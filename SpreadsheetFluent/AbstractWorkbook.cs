using OfficeOpenXml;
using SpreadsheetFluent.Blocks;
using SpreadsheetFluent.Enums;
using SpreadsheetFluent.Extensions;
using SpreadsheetFluent.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
namespace SpreadsheetFluent
{

    public abstract class AbstractWorkbook : IWorkBookCreator
    {
        private BlockPropertyLocations _blockPropertyLocation;
        private BlockPropertyLocations BlockPropertyLocation
        {
            get
            {
                if (_blockPropertyLocation == null)
                    _blockPropertyLocation = new BlockPropertyLocations();
                return _blockPropertyLocation;
            }
        }
        internal IList<AbstractWorksheet> Worksheets { get; } = new List<AbstractWorksheet>();
        internal Action<ExcelWorkbook> ExcelWorkbookOptions { get; private set; }
        protected AbstractWorkbook AddWorksheet<T>() where T : AbstractWorksheet, new()
        {
            Worksheets.Add(new T());
            return this;
        }
        protected AbstractWorkbook AddWorksheet(AbstractWorksheet worksheet)
        {
            Worksheets.Add(worksheet);
            return this;
        }

        protected AbstractWorkbook Configure(Action<ExcelWorkbook> opt)
        {
            ExcelWorkbookOptions = opt;
            return this;
        }

        public byte[] Generate(params IEnumerable<object>[] datasets)
        {
            using (var packages = CreateExcel(datasets))
            {
                return packages.GetAsByteArray();
            }
        }

        public ExcelPackage CreateExcel(params IEnumerable<object>[] datasets)
        {
            var packages = new ExcelPackage();
            foreach (var wsc in Worksheets)
            {
                var worksheet = packages.Workbook.Worksheets.Any(p => p.Name == wsc.Name)
                                        ? packages.Workbook.Worksheets[wsc.Name]
                                        : packages.Workbook.Worksheets.Add(wsc.Name);


                (int Row, int Column) lastPosition = (0, 0);
                for (int index = 0; index < datasets.Length; index++)
                {
                    var dataset = datasets[index];
                    var datasetType = dataset.GetType();

                    var datasetArgumentType = datasetType.IsGenericType
                                                ? datasetType.GetGenericArguments()[0]
                                                : datasetType.GetElementType();

                    var blocksStatic = wsc.Blocks.Where(p => p.BlockType == datasetArgumentType);
                    if (blocksStatic == null || !blocksStatic.Any()) continue;

                    foreach (var blockStatic in blocksStatic)
                    {
                        var block = blockStatic.Clone<WorksheetBlockBase>();

                        if (blockStatic.Direction == BlockDirection.Above) block.StartRow += lastPosition.Row;
                        else block.StartColumn += lastPosition.Column;

                        lastPosition = WriteHeader(block, ref worksheet);
                        lastPosition = WriteBody(dataset, block, ref worksheet);

                        var allRange = worksheet.Cells[blockStatic.StartRow, blockStatic.StartColumn, lastPosition.Row, lastPosition.Column];
                        block.ApplyStyle(WorksheetSection.All, ref allRange);
                    }
                }
                wsc.ExcelWorksheetOptions?.Invoke(worksheet);
            }
            return packages;
        }

        private (int Row, int Column) WriteHeader(WorksheetBlockBase block, ref ExcelWorksheet worksheet)
        {
            var propertyLocations = new List<PropertyLocation>();
            //Write title
            var lastColumnHeaderPos = block.StartColumn + block.Rules.Count - 1;
            if (!string.IsNullOrWhiteSpace(block.Title?.Value))
            {
                var titleRow = block.StartRow;
                worksheet.Cells[titleRow, block.StartColumn].Value = block.Title.Value;
                var cellRange = worksheet.Cells[titleRow, block.StartColumn, titleRow, lastColumnHeaderPos];
                cellRange.Merge = true;

                block.ApplyStyle(WorksheetSection.Title, ref cellRange);
                block.ApplyStyle(WorksheetSection.EachCell, ref cellRange);
                block.StartRow++;
            }

            //Write headers
            for (int index = 0; index < block.Rules.Count; index++)
            {

                var rule = block.Rules[index];
                var column = block.StartColumn + index;
                var cellRange = worksheet.Cells[block.StartRow, column];
                cellRange.Value = rule.Caption?.Value ?? rule.Expression.GetMember().Name;

                rule.AutoFit?.Invoke(cellRange);

                rule.ApplyStyle(WorksheetSection.Caption, ref cellRange);
                block.ApplyStyle(WorksheetSection.EachCell, ref cellRange);

                propertyLocations.Add(new PropertyLocation(PropertyLocation.PropertyLocationSectionType.Header, rule.GetHashCode(), block.StartRow, column));
            }
            var header = worksheet.Cells[block.StartRow, block.StartColumn, block.StartRow, lastColumnHeaderPos];
            block.ApplyStyle(WorksheetSection.Header, ref header);

            BlockPropertyLocation.AddRange(block.Id, propertyLocations);

            return (block.StartRow, lastColumnHeaderPos);
        }

        private (int Row, int Column) WriteBody(IEnumerable<object> dataset, WorksheetBlockBase block, ref ExcelWorksheet worksheet)
        {
            var currentRow = block.StartRow;
            var currentColumn = 0;
            var propertyLocations = new List<PropertyLocation>();
            foreach (var data in dataset)
            {
                currentRow++;
                for (int index = 0; index < block.Rules.Count; index++)
                {
                    currentColumn = block.StartColumn + index;
                    var rule = block.Rules[index];
                    var cellRange = worksheet.Cells[currentRow, currentColumn];
                    cellRange.Value = rule.GetValue(data);
                    rule.AutoFit?.Invoke(cellRange);
                    rule.ApplyStyle(WorksheetSection.Cell, ref cellRange);
                    rule.ApplyStyle(WorksheetSection.EachCell, ref cellRange);

                    rule.Options?.Invoke(cellRange);
                    propertyLocations.Add(new PropertyLocation(PropertyLocation.PropertyLocationSectionType.Body, rule.GetHashCode(), currentRow, currentColumn));
                }

                var bodyRange = worksheet.Cells[block.StartRow, block.StartColumn, currentRow, currentColumn];
                block.ApplyStyle(WorksheetSection.Body, ref bodyRange);
            }

            BlockPropertyLocation.AddRange(block.Id, propertyLocations);

            return (currentRow, currentColumn);
        }
    }
}
