using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelManipulator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
                 backgroundWorker1.WorkerReportsProgress = true;
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


                progressBar1.Visible = false;
                progressBar1.Value = 0;
                lblStatus.Text = "";

            }
        }


        // Este es el método que va a ejecutar el BackgroundWorker
        private void btnProcessFile_Click(object sender, EventArgs e)
        {


            // Muestra la barra de progreso y el mensaje
            progressBar1.Visible = true;
            lblStatus.Visible = true;
            lblStatus.Text = "Iniciando el proceso...";

            string filePath = txtFilePath.Text;
            string startDate = dtpStartDate.Value.ToString("yyyy-MM-dd"); // Formato de fecha
            string endDate = dtpEndDate.Value.ToString("yyyy-MM-dd");



          


            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Por favor, selecciona un archivo Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            backgroundWorker1.RunWorkerAsync(new string[] { filePath, startDate, endDate });

        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            // Obtener los parámetros de entrada
            string[] args = (string[])e.Argument;
            string filePath = args[0];
            string startDate = args[1];
            string endDate = args[2];

            // Inicializar la barra de progreso
            progressBar1.Maximum = 100;

            try
            {
                // Crear el duplicador de Excel
                ExcelDuplicator duplicator = new ExcelDuplicator();
                // Pasar el BackgroundWorker para actualizar el progreso

                duplicator.ProcessExcelFile(filePath, startDate, endDate, backgroundWorker1);
               
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Este método se llama cuando el BackgroundWorker reporta progreso
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Actualizar la barra de progreso

            progressBar1.Value = e.ProgressPercentage;

            // Actualiza el mensaje de estado
            switch (e.ProgressPercentage)
            {
                case 20:
                    lblStatus.Text = "Duplicando archivos...";
                    break;
                case 50:
                    lblStatus.Text = "Editando campos...";
                    break;
                case 70:
                    lblStatus.Text = "Guardando cambios...";
                    break;
                case 90:
                    lblStatus.Text = "Cerrando archivo...";
                    break;
                case 100:
                    lblStatus.Text = "Proceso finalizado.";
                    Task.Delay(2000);
                 
                    break;
            }
        }

    
        // Este método se llama cuando el BackgroundWorker termina
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                // Si hay un error, mostrar mensaje
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // Cuando termina, mostrar mensaje de éxito
                MessageBox.Show("El archivo Excel se procesó correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


    }
}

