using MySql.Data.MySqlClient;
using OfficeOpenXml;

namespace ExcelManipulator
{

    public class bbdd
    {
        private string connectionString = "Server=localhost;Database=ExcelRojas;User=excel2;Password=1234;";

        /// <param name="filePath">Ruta del archivo Excel.</param>
        public void SaveExcelDataToDatabase(string filePath)
        {
            // Configurar el contexto de licencia de EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                // Obtener las hojas
                var sheet1 = package.Workbook.Worksheets[0];
                var sheet2 = package.Workbook.Worksheets.Count > 1 ? package.Workbook.Worksheets[1] : null;

                // Guardar los datos de la primera hoja en la tabla precios_compra
                if (sheet1 != null)
                {
                    SaveSheet1ToDatabase(sheet1);
                }

                // Guardar los datos de la segunda hoja en la tabla centros_ocio
                if (sheet2 != null)
                {
                    SaveSheet2ToDatabase(sheet2);
                }
            }
        }

        /// <param name="sheet1">Hoja de Excel.</param>
        private void SaveSheet1ToDatabase(ExcelWorksheet sheet1)
        {
            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 4; row <= sheet1.Dimension.End.Row; row++)
                {
                    var command = new MySqlCommand(
                        "INSERT INTO precios_compra (num_producto, num_proveedor, fecha_inicio, codigo_divisa, codigo_variante, " +
                        "codigo_unidad_medida, cantidad_minima, num_linea, fecha_final, coste_unidad, descuento, " +
                        "otros_impuestos, rappel, para_todos, fecha_creacion, id_usuario, id_interno, num_referencia) " +
                        "VALUES (@num_producto, @num_proveedor, @fecha_inicio, @codigo_divisa, @codigo_variante, " +
                        "@codigo_unidad_medida, @cantidad_minima, @num_linea, @fecha_final, @coste_unidad, @descuento, " +
                        "@otros_impuestos, @rappel, @para_todos, @fecha_creacion, @id_usuario, @id_interno, @num_referencia) " +
                        "ON DUPLICATE KEY UPDATE num_producto = VALUES(num_producto), num_proveedor = VALUES(num_proveedor);",
                        connection);

                    command.Parameters.AddWithValue("@num_producto", sheet1.Cells[row, 1].Text);
                    command.Parameters.AddWithValue("@num_proveedor", sheet1.Cells[row, 2].Text);
                    command.Parameters.AddWithValue("@fecha_inicio", sheet1.Cells[row, 3].Text);
                    command.Parameters.AddWithValue("@codigo_divisa", sheet1.Cells[row, 4].Text);
                    command.Parameters.AddWithValue("@codigo_variante", sheet1.Cells[row, 5].Text);
                    command.Parameters.AddWithValue("@codigo_unidad_medida", sheet1.Cells[row, 6].Text);
                    command.Parameters.AddWithValue("@cantidad_minima", sheet1.Cells[row, 7].Text);
                    command.Parameters.AddWithValue("@num_linea", sheet1.Cells[row, 8].Text);
                    command.Parameters.AddWithValue("@fecha_final", sheet1.Cells[row, 9].Text);
                    command.Parameters.AddWithValue("@coste_unidad", sheet1.Cells[row, 10].Text);
                    command.Parameters.AddWithValue("@descuento", sheet1.Cells[row, 11].Text);
                    command.Parameters.AddWithValue("@otros_impuestos", sheet1.Cells[row, 12].Text);
                    command.Parameters.AddWithValue("@rappel", sheet1.Cells[row, 13].Text);
                    command.Parameters.AddWithValue("@para_todos", sheet1.Cells[row, 14].Text);
                    command.Parameters.AddWithValue("@fecha_creacion", sheet1.Cells[row, 15].Text);
                    command.Parameters.AddWithValue("@id_usuario", sheet1.Cells[row, 16].Text);
                    command.Parameters.AddWithValue("@id_interno", sheet1.Cells[row, 17].Text);
                    command.Parameters.AddWithValue("@num_referencia", sheet1.Cells[row, 18].Text);

                    command.ExecuteNonQuery();
                }
            }
        }


        /// <param name="sheet2">Hoja de Excel.</param>
        private void SaveSheet2ToDatabase(ExcelWorksheet sheet2)
        {
            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 4; row <= sheet2.Dimension.End.Row; row++)
                {
                    var command = new MySqlCommand(
                        "INSERT INTO centros_ocio (id_interno, num_cliente, cod_direccion) " +
                        "VALUES (@id_interno, @num_cliente, @cod_direccion) " +
                        "ON DUPLICATE KEY UPDATE num_cliente = VALUES(num_cliente), cod_direccion = VALUES(cod_direccion);",
                        connection);

                    command.Parameters.AddWithValue("@id_interno", sheet2.Cells[row, 1].Text);
                    command.Parameters.AddWithValue("@num_cliente", sheet2.Cells[row, 2].Text);
                    command.Parameters.AddWithValue("@cod_direccion", sheet2.Cells[row, 3].Text);

                    command.ExecuteNonQuery();
                }
            }
        }
    }
}