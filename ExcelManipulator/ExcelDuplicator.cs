using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using ExcelManipulator;
using System.ComponentModel;
using Mysqlx.Crud;

public class ExcelDuplicator
{
    private int CurrentRow;
    private int NumRowsPreciosCompra;
    private int NumRowsComercioOcio;
    private ExcelWorksheet sheet1;
    private ExcelWorksheet sheet2;
    private int progress = 0;
    private int pasostotales;
    public void ProcessExcelFile(string filePath, string startDate, string endDate, BackgroundWorker worker)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        // Abrir el archivo Excel
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            // Acceder a las hojas
            this.sheet1 = package.Workbook.Worksheets[0];
            this.sheet2 = package.Workbook.Worksheets.Count > 1 ? package.Workbook.Worksheets[1] : null;

            //Declarar cuantas filas hay en cada hoja
            this.NumRowsPreciosCompra = GetNumRows(this.sheet1);
            this.NumRowsComercioOcio = GetNumRows(this.sheet2);

            //Contar todas los pasos que se han de hacer

            pasostotales = CeldasTotales();

            // Obtener el ID más alto de la hoja "precios compra"
            int maxId = GetMaxId(4, 17);
            //worker.ReportProgress(20);

            // Duplicar registros en la primera hoja "precios compra"
            DuplicarRegistrosHojaPreciosCompra(maxId, startDate, worker);
            //worker.ReportProgress(40);

            // Asignar fecha_fin a las celdas sin fecha
            AsignarFechaFinPreciosCompra(4, 9, endDate, worker);
            //worker.ReportProgress(60);

            // Duplicar registros en la segunda hoja (si existe) 
            // por cada "id interno" de Precios compra 
            if (sheet2 != null)
            {
                DuplicarRegistrosHojaComerciosOcio(maxId, startDate, worker);
              //  worker.ReportProgress(80);
            }

            // Guardar cambios en un nuevo archivo
            GuardarExcel(package, filePath,worker);
          //  worker.ReportProgress(100);
        }
    }

    private int GetMaxId( int startRow, int idColumn)
    {
        int lastRow = this.sheet1.Dimension.End.Row;
        int maxId = 0;

        for (int row = startRow; row <= lastRow; row++)
        {
            if (int.TryParse(this.sheet1.Cells[row, idColumn].Text, out int id))
            {
                maxId = Math.Max(maxId, id);
            }
        }

        return maxId;
    }

    //Devolver el numero total de filas a partir de una inicial
    private int GetNumRows(ExcelWorksheet sheet) 
    {
       return sheet.Dimension.End.Row;
    }

    private void DuplicarRegistrosHojaPreciosCompra( int maxId, string startDate, BackgroundWorker worker)
    {
        int lastRow = this.sheet1.Dimension.End.Row;

        for (this.CurrentRow = 4; this.CurrentRow <= lastRow; this.CurrentRow++)
        {
            int newRow = lastRow + (this.CurrentRow - 3);

            // Copiar fila y cambiar color de fondo
            for (int col = 1; col <= this.sheet1.Dimension.End.Column; col++)
            {
                this.sheet1.Cells[newRow, col].Value = this.sheet1.Cells[this.CurrentRow, col].Value;
                this.sheet1.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                this.sheet1.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);


                this.progress++;
                worker.ReportProgress((this.progress*100)/ pasostotales);
            }

            // Asignar nuevo ID y fecha de inicio
            int suma = Convert.ToInt32(this.sheet1.Cells[newRow, 17].Value);
            this.sheet1.Cells[newRow, 17].Value = (maxId + suma).ToString(); // ID con 8 dígitos
            this.sheet1.Cells[newRow, 3].Value = startDate; // Columna fecha_inicio
        }
    }

    private void DuplicarRegistrosHojaComerciosOcio(int maxId, string startDate, BackgroundWorker worker)
    {
        int lastRow = this.sheet2.Dimension.End.Row;

        for (this.CurrentRow = 4; this.CurrentRow <= lastRow; this.CurrentRow++)
        {
            int newRow = lastRow + (this.CurrentRow - 3);

            // Copiar fila y cambiar color de fondo
            for (int col = 1; col <= this.sheet2.Dimension.End.Column; col++)
            {
                this.sheet2.Cells[newRow, col].Value = this.sheet2.Cells[this.CurrentRow, col].Value;
                this.sheet2.Cells[newRow, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                this.sheet2.Cells[newRow, col].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                this.progress++;
                worker.ReportProgress((this.progress * 100) / pasostotales);
            }

            // Asignar nuevo ID y fecha de inicio
            int suma = Convert.ToInt32(this.sheet2.Cells[newRow, 1].Value);
            this.sheet2.Cells[newRow, 1].Value = (maxId + suma).ToString(); // ID
        }
    }

    private void AsignarFechaFinPreciosCompra( int startRow, int dateColumn, string endDate, BackgroundWorker worker)
    {
        int lastRow = this.sheet1.Dimension.End.Row;

        for (this.CurrentRow = startRow; this.CurrentRow <= lastRow; this.CurrentRow++)
        {
            if (string.IsNullOrWhiteSpace(this.sheet1.Cells[this.CurrentRow, dateColumn].Text))
            {
                this.sheet1.Cells[this.CurrentRow, dateColumn].Value = endDate;
                this.sheet1.Cells[this.CurrentRow, dateColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                this.sheet1.Cells[this.CurrentRow, dateColumn].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                this.progress++;
                worker.ReportProgress((this.progress * 100) / pasostotales);
            }
        }
    }

    private void GuardarExcel(ExcelPackage package, string filePath, BackgroundWorker worker)
    {

        worker.ReportProgress(100);
        string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), "Modificado_" + Path.GetFileName(filePath));
        package.SaveAs(new FileInfo(newFilePath));

        // Guardado en base de datos, proxima feature, ya funciona
       // bbdd dbHandler = new bbdd();
       // dbHandler.SaveExcelDataToDatabase(filePath);
    }



    //Calcula todas las celdas a editar, Todas las de "Precios Compra" + todas "Comercio Ocio" + las de la columna de "id_interno" en "Precios compra" y "Fecha_fin" en precios compra
    public int CeldasTotales()
    {
            int Hoja1 = sheet1.Dimension.End.Row * sheet1.Dimension.End.Column;
            int Hoja2 = sheet2.Dimension.End.Row * sheet2.Dimension.End.Column;



            return Hoja1 + Hoja2 + (sheet1.Dimension.End.Row*2);
    }

}

