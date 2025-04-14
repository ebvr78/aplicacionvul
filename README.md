# aplicacionvul
using System;
using System.Data;
using System.Data.SQLite;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace RegistroProductosApp
{
    public partial class MainForm : Form
    {
        private string connectionString = "Data Source=productos.db;Version=3;";

        public MainForm()
        {
            InitializeComponent();
            CrearBaseDeDatos();
            CargarDatos();
        }

        private void CrearBaseDeDatos()
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                var command = new SQLiteCommand(
                    "CREATE TABLE IF NOT EXISTS Productos (" +
                    "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                    "Nombre TEXT NOT NULL, " +
                    "Fardos INTEGER NOT NULL, " +
                    "Unidades INTEGER NOT NULL, " +
                    "Fecha TEXT NOT NULL)", connection);
                command.ExecuteNonQuery();
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtProducto.Text))
            {
                MessageBox.Show("Ingresa el nombre del producto.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                var command = new SQLiteCommand(
                    "INSERT INTO Productos (Nombre, Fardos, Unidades, Fecha) VALUES (@nombre, @fardos, @unidades, @fecha)",
                    connection);
                command.Parameters.AddWithValue("@nombre", txtProducto.Text);
                command.Parameters.AddWithValue("@fardos", numFardos.Value);
                command.Parameters.AddWithValue("@unidades", numUnidades.Value);
                command.Parameters.AddWithValue("@fecha", dateTimePicker.Value.ToString("yyyy-MM-dd"));
                command.ExecuteNonQuery();
            }

            MessageBox.Show("Producto registrado correctamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LimpiarCampos();
            CargarDatos();
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Productos");
                    
                    worksheet.Cell(1, 1).Value = "Producto";
                    worksheet.Cell(1, 2).Value = "Fardos";
                    worksheet.Cell(1, 3).Value = "Unidades";
                    worksheet.Cell(1, 4).Value = "Fecha";

                    using (var connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        var adapter = new SQLiteDataAdapter("SELECT Nombre, Fardos, Unidades, Fecha FROM Productos", connection);
                        var table = new DataTable();
                        adapter.Fill(table);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            worksheet.Cell(i + 2, 1).Value = table.Rows[i]["Nombre"].ToString();
                            worksheet.Cell(i + 2, 2).Value = table.Rows[i]["Fardos"].ToString();
                            worksheet.Cell(i + 2, 3).Value = table.Rows[i]["Unidades"].ToString();
                            worksheet.Cell(i + 2, 4).Value = table.Rows[i]["Fecha"].ToString();
                        }
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Files|*.xlsx",…
