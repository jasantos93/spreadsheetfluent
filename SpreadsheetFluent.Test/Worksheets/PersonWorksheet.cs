using SpreadsheetFluent.Test.Models;

namespace SpreadsheetFluent.Test.Worksheets
{
    public class PersonInfoWorksheet : AbstractWorksheet
    {
        public PersonInfoWorksheet()
        {

            CreateBlock<Person>()
                .WithTitle("Basico")
                .WithStyle()
                    .ForHeader(opt =>
                    {
                        opt.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.MediumDashDot);
                        opt.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        opt.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);

                    })
                    .ForAll(opt =>
                    {
                        opt.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Double);
                        opt.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        opt.Fill.BackgroundColor.SetColor(System.Drawing.Color.GhostWhite);
                    })
                    .ForTitle(opt =>
                    {
                        opt.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        opt.WrapText = true;
                        opt.Font.Size = 20;
                        opt.Font.Bold = true;
                    })
                .WithColumn(p => p.Name)
                    .AutoFit()
                    .WithCaption("Nombre")
                    .WithStyle()
                        .ForCaption(opt => opt.Font.Bold = true)
                .WithColumn(p => p.Birthday.DateTime)
                    .WithCaption("Fecha de nacimiento")
                    .WithStyle()
                        .ForCell(opt => opt.Numberformat.Format = "m/d/yyyy h:mm")
                .WithColumn(p => p.Age)
                    .WithCaption("Edad")
                    .Configure(opt =>
                    {
                        if (opt.GetValue<int>() > 150)
                            opt.AddComment("Who lives so much?", "Machine");
                    })
                    .WithStyle()
                        .ForCaption(opt => opt.Font.Italic = true);

            CreateBlock<Person>()
                .WithStartRow(1)
                .WithStyle()
                    .ForHeader(opt =>
                    {
                        opt.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dashed);
                        opt.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        opt.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                    })
                .WithColumn(p => p.Salary);

            CreateBlock<Department>()
                .WithDirection(Enums.BlockDirection.Above)
                .WithTitle("Departmento")
                .WithStartColumn(1)
                .WithColumn(p => p.Name)
                .WithColumn(p => p.Description)
                    .AutoFit();

        }
    }
}
