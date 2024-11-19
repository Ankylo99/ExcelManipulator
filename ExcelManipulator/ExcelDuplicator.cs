using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;

public class ExcelDuplicator
{
    public void ProcessExcelFile(string filePath, String startDate, String endDate)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        // Abrir el archivo Excel
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            // Acceder a la primera hoja
            var sheet1 = package.Workbook.Worksheets[0];
            var sheet2 = package.Workbook.Worksheets.Count > 1 ? package.Workbook.Worksheets[1] : null;

            // 1. Obtener el ID más alto
            int lastRow = sheet1.Dimension.End.Row;
            int maxId = 0;
            for (int row = 2; row <= lastRow; row++) // Suponiendo que la fila 1 es el encabezado
            {
                int id = Convert.ToInt32(sheet1.Cells[row, 1].Text);
                if (id > maxId) maxId = id;
            }

            // 2. Duplicar registros y 3. Asignar fecha_inicio
            for (int row = 2; row <= lastRow; row++)
            {
                maxId++; // Incrementar ID
                int newRow = lastRow + (row - 1);

                // Copiar fila y cambiar color de fondo de celdas
                for (int col = 1; col <= sheet1.Dimension.End.Column; col++)
                {
                    sheet1.Cells[newRow, col].Value = sheet1.Cells[row, col].Value;
                    sheet1.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet1.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                }

                // Asignar nuevo ID y fecha de inicio
                sheet1.Cells[newRow, 1].Value = maxId; // Columna ID
                sheet1.Cells[newRow, 2].Value = startDate; // Columna fecha_inicio
     
            }

            // 4. Asignar fecha_fin a las celdas sin fecha en esa columna
            for (int row = 2; row <= sheet1.Dimension.End.Row; row++)
            {
                if (string.IsNullOrWhiteSpace(sheet1.Cells[row, 3].Text)) // Columna fecha_fin
                {
                    sheet1.Cells[row, 3].Value = endDate;
                    sheet1.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet1.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                }
            }

            // 7. Duplicar registros en la segunda hoja, si existe
            if (sheet2 != null)
            {
                int sheet2LastRow = sheet2.Dimension.End.Row;
                int sheet2NewRowStart = sheet2LastRow + 1;

                for (int row = 2; row <= lastRow; row++)
                {
                    int newRow = sheet2NewRowStart + (row - 2);
                    for (int col = 1; col <= sheet2.Dimension.End.Column; col++)
                    {
                        sheet2.Cells[newRow, col].Value = sheet1.Cells[row, col].Value;
                        sheet2.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet2.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                }
            }

            // 8. Guardar cambios en un nuevo archivo
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), "Modified_" + Path.GetFileName(filePath));
            package.SaveAs(new FileInfo(newFilePath));
        }
    }
}