using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using ExcelManipulator;

public class ExcelDuplicator
{
    public void ProcessExcelFile(string filePath, string startDate, string endDate)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        // Abrir el archivo Excel
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            // Acceder a las hojas
            var sheet1 = package.Workbook.Worksheets[0];
            var sheet2 = package.Workbook.Worksheets.Count > 1 ? package.Workbook.Worksheets[1] : null;

            // Obtener el ID más alto
            int maxId = GetMaxId(sheet1, 4, 17);

            // Duplicar registros en la primera hoja
            DuplicarFilas1(sheet1, maxId, startDate);

            // Asignar fecha_fin a las celdas sin fecha
            FechaFin(sheet1, 4, 9, endDate);

            // Duplicar registros en la segunda hoja (si existe)
            if (sheet2 != null)
            {
                DuplicarFilas2(sheet2, maxId, startDate);
            }

            // Guardar cambios en un nuevo archivo
            GuardarExcel(package, filePath);
        }
    }

    private int GetMaxId(ExcelWorksheet sheet, int startRow, int idColumn)
    {
        int lastRow = sheet.Dimension.End.Row;
        int maxId = 0;

        for (int row = startRow; row <= lastRow; row++)
        {
            if (int.TryParse(sheet.Cells[row, idColumn].Text, out int id))
            {
                maxId = Math.Max(maxId, id);
            }
        }

        return maxId;
    }

    private void DuplicarFilas1(ExcelWorksheet sheet, int maxId, string startDate)
    {
        int lastRow = sheet.Dimension.End.Row;

        for (int row = 4; row <= lastRow; row++)
        {
            int newRow = lastRow + (row - 3);

            // Copiar fila y cambiar color de fondo
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                sheet.Cells[newRow, col].Value = sheet.Cells[row, col].Value;
                sheet.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }

            // Asignar nuevo ID y fecha de inicio
            int suma = Convert.ToInt32(sheet.Cells[newRow, 17].Value);
            sheet.Cells[newRow, 17].Value = (maxId + suma).ToString(); // ID con 8 dígitos
            sheet.Cells[newRow, 3].Value = startDate; // Columna fecha_inicio
        }
    }

    private void DuplicarFilas2(ExcelWorksheet sheet, int maxId, string startDate)
    {
        int lastRow = sheet.Dimension.End.Row;

        for (int row = 4; row <= lastRow; row++)
        {
            int newRow = lastRow + (row - 3);

            // Copiar fila y cambiar color de fondo
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                sheet.Cells[newRow, col].Value = sheet.Cells[row, col].Value;
                sheet.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }

            // Asignar nuevo ID y fecha de inicio
            int suma = Convert.ToInt32(sheet.Cells[newRow, 1].Value);
            sheet.Cells[newRow, 1].Value = (maxId + suma).ToString(); // ID
        }
    }

    private void FechaFin(ExcelWorksheet sheet, int startRow, int dateColumn, string endDate)
    {
        int lastRow = sheet.Dimension.End.Row;

        for (int row = startRow; row <= lastRow; row++)
        {
            if (string.IsNullOrWhiteSpace(sheet.Cells[row, dateColumn].Text))
            {
                sheet.Cells[row, dateColumn].Value = endDate;
                sheet.Cells[row, dateColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[row, dateColumn].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }
        }
    }

    private void GuardarExcel(ExcelPackage package, string filePath)
    {
        string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), "Modified_" + Path.GetFileName(filePath));
        package.SaveAs(new FileInfo(newFilePath));

        bbdd dbHandler = new bbdd();
        dbHandler.SaveExcelDataToDatabase(filePath);
    }
}
