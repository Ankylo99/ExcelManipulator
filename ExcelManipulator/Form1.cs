using OfficeOpenXml.Table;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelManipulator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private async void btnProcessFile_ClickAsync(object sender, EventArgs e)
        {
            loadingIcon.Visible = true;
            statusLabel.Text = "Cargando archivo...";

            

            string filePath = txtFilePath.Text;
            string startDate = dtpStartDate.Value.ToString("yyyy-MM-dd"); // Formato de fecha
            string endDate = dtpEndDate.Value.ToString("yyyy-MM-dd");

            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Por favor, selecciona un archivo Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
               
            }

            try
            {
                await Task.Run(() =>
                {
                    var duplicator = new ExcelDuplicator(statusLabel);
                    duplicator.ProcessExcelFile(filePath, startDate, endDate);
                    MessageBox.Show("El archivo Excel se procesó correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                loadingIcon.Visible = false;
                statusLabel.Text = "";
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}

