using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace SParametersExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void btnExcelOpen_Click(object sender, EventArgs e)
        {
            string filePath = DosyaAc();
            groupBox1.Text = filePath;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DataTable dataTable = new DataTable();
                    dataTable = ExcelDosyasiniOku(filePath);
                    
                    dataGridView1.DataSource = dataTable;
                    dataGridViewSutunBaslikAyar(dataTable, 6);                  // 6 sütun sayisi
                    btnSorgula.Enabled = true;
                    //VerileriGrafikteGoster(dataTable);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu:" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            comboBoxItemAdd(filePath);
            comboBoxSheetNames.SelectedIndex = 0;
        }

        private void dataGridViewSutunBaslikAyar(DataTable dataTable, int sonSutunSayisi)
        {
            if (dataTable.Rows.Count > 2)
            {
                for (int i = 0; i < sonSutunSayisi; i++)
                {
                    dataGridView1.Columns[i].HeaderText = dataTable.Rows[1][i].ToString();
                }
                //dataGridView1.Rows.Remove();
            }
        }
        private void comboBoxItemAdd(string filePath)
        {
            comboBoxSheetNames.Items.Clear();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    comboBoxSheetNames.Items.Add(sheet);
                }
            }
        }
        private string DosyaAc()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.ShowDialog();
            return openFileDialog.FileName;
        }
        private void comboBoxSheetNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = groupBox1.Text;
            string selectedSheet = comboBoxSheetNames.SelectedItem.ToString();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[selectedSheet];
                if (worksheet != null)
                {
                    //Sütun satırı belirleme
                    int startRow = 1;
                    //Sütun sayisi
                    int sonSutunSayisi = worksheet.Dimension.End.Column;
                    sonSutunSayisi = 6;
                    int sonSatirSayisi = worksheet.Dimension.End.Row;

                    DataTable dataTable = new DataTable();

                    foreach (var firstRowCell in worksheet.Cells[startRow, 1, startRow, sonSutunSayisi])
                    {
                        if (double.TryParse(firstRowCell.Text, out double numericValue))
                        {
                            dataTable.Columns.Add(firstRowCell.Text, typeof(double));
                        }
                        else
                        {
                            dataTable.Columns.Add(firstRowCell.Text);
                        }
                    }

                    for (var rowNumber = (startRow + 1); rowNumber <= sonSatirSayisi; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, sonSutunSayisi];
                        var newRow = dataTable.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }
                        dataTable.Rows.Add(newRow);
                    }
                    dataGridView1.DataSource = dataTable;
                    //dataGridViewSutunBaslikAyar(dataTable, 6);
                    //VerileriGrafikteGoster(dataTable);
                }
            }
            //dataGridView1.DataSource = ExcelDosyasiniOku(filePath); //kullanabilmek için selectedSheet için düzenlenmeli
        }

        private void comboBoxAralikMin_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxAralikMax.Enabled = true;
            comboBoxAralikMax.Items.Clear();

            int minSecimi = int.Parse(comboBoxAralikMin.SelectedItem.ToString());

            comboBoxAralikMax.Items.Clear();

            for (int i = (minSecimi + 500); i <= 4000; i += 500)
            {
                comboBoxAralikMax.Items.Add(i.ToString());
            }
            comboBoxAralikMax.Text = comboBoxAralikMin.SelectedItem.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxAralikMin.Items.AddRange(new string[] { "300", "500", "1000", "1500", "2000", "2500", "3000", "3500", "4000" });
            comboBoxAralikMax.Enabled = false;
            btnSorgula.Enabled = false;
            btnSave.Enabled = false;
        }

        private void btnSorgula_Click(object sender, EventArgs e)
        {
            string selectedMinValue = comboBoxAralikMin.SelectedItem?.ToString();

            double minMHz = string.IsNullOrEmpty(selectedMinValue) ? 0 : double.Parse(selectedMinValue);
            double maxMHz = double.Parse(comboBoxAralikMax.Text.ToString());

            //DataTable updatedData = ExcelDosyasiniOku(groupBox1.Text);
            // Bu işlem sayfalar arası sorgulamada hata verdiriyor. Misal 4. sayfayı sorgulama yaparken dosya yeniden okunduğu için varsayılan olarak 1. sayfayı çekiyor 1. sayfa üzerinden sorgu yapıyor.

            DataGridViewColumn mhzColumn = dataGridView1.Columns.Cast<DataGridViewColumn>().FirstOrDefault(col => col.HeaderText == "MHz");

            if (mhzColumn != null)
            {
                DataTable originalData = (DataTable)dataGridView1.DataSource;

                DataTable filteredData = FilterByMHz(originalData, minMHz, maxMHz);

                dataGridView1.DataSource = filteredData;

                btnSave.Enabled = true;

                VerileriGrafikteGoster(filteredData);

            }
            else
            {
                MessageBox.Show("Sütun başlığı 'MHz' olan bir sütun bulunamadı", "Hata:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable FilterByMHz(DataTable originalTable, double minMHz, double maxMHz)
        {
            DataTable filteredTable = originalTable.Clone();

            string mhzColumnName = "Column1";

            Type columnType = null;

            foreach (DataColumn column in originalTable.Columns)
            {
                if (column.ColumnName.Equals(mhzColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnType = column.DataType;
                    break;
                }
            }

            if (columnType == null)
            {
                MessageBox.Show("Sütun başlığı 'MHz' olan bir sütun bulunamadı", "Hata:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return filteredTable;
            }

            foreach (DataRow row in originalTable.Rows)
            {
                if (columnType == typeof(string))
                {
                    string mhzValue = row[mhzColumnName].ToString();
                    if (Double.TryParse(mhzValue, out double MHz))
                    {
                        if (MHz >= minMHz && MHz <= maxMHz)
                        {
                            filteredTable.ImportRow(row);
                        }
                    }
                }
                else if (columnType == typeof(double))
                {
                    double MHz = Convert.ToDouble(row[mhzColumnName]);
                    if (MHz >= minMHz && MHz <= maxMHz)
                    {
                        filteredTable.ImportRow(row);
                    }
                }
            }
            return filteredTable;
        }

        private DataTable ExcelDosyasiniOku(string filePath)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    int startRow = 1;
                    int sonSutunSayisi = worksheet.Dimension.End.Column;
                    sonSutunSayisi = 6;
                    int sonSatirSayisi = worksheet.Dimension.End.Row;

                    for (int columnIndex = 1; columnIndex <= sonSutunSayisi; columnIndex++)
                    {
                        string columnHeader = worksheet.Cells[startRow, columnIndex].Text;
                        //DataColumn column = new DataColumn(columnHeader,typeof(double));
                        dataTable.Columns.Add(columnHeader);
                    }

                    for (var rowNumber = (startRow + 1); rowNumber <= sonSatirSayisi; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, sonSutunSayisi];
                        var newRow = dataTable.NewRow();
                        foreach (var cell in row)
                        {
                            if (double.TryParse(cell.Text, out double numericValue))
                            {
                                newRow[cell.Start.Column - 1] = numericValue;
                            }
                            else
                            {
                                newRow[cell.Start.Column - 1] = cell.Text;
                            }
                        }
                        dataTable.Rows.Add(newRow);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dataTable;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string saveSheetName = textBoxSaveName.Text;

            if (string.IsNullOrEmpty(saveSheetName))
            {
                MessageBox.Show("Sayda ad giriniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double minMHz = double.Parse(comboBoxAralikMin.SelectedItem.ToString());
            double maxMHz = double.Parse(comboBoxAralikMax.Text.ToString());

            DataTable originalData = (DataTable)dataGridView1.DataSource;
            DataTable filteredData = FilterByMHz(originalData, minMHz, maxMHz);


            using (var package = new ExcelPackage(new FileInfo(groupBox1.Text)))
            {
                ExcelWorksheet exitisingWorksheet = package.Workbook.Worksheets[saveSheetName];

                if (exitisingWorksheet != null)
                {
                    DialogResult result = MessageBox.Show("Girilen sayfa adında bir sayfa zaten var. Verileri üzerine yazmak istiyor musunuz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        package.Workbook.Worksheets.Delete(exitisingWorksheet);
                        SaveFilteredDataToWorksheet(package, filteredData, saveSheetName);

                        package.Save();
                        MessageBox.Show("Sorgulanmış veriler sayfaya kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("veriler sayfaya kaydedilmedi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    SaveFilteredDataToWorksheet(package, filteredData, saveSheetName);
                    SaveGrafik(package, chart1, saveSheetName);
                    package.Save();
                    MessageBox.Show("Sorgulanmış veriler yeni sayfaya kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }
        private void SaveFilteredDataToWorksheet(ExcelPackage package, DataTable filteredData, string sheetName)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName); // Yeni bir sayfa oluştur

            // Başlık satırını ekle
            for (int col = 0; col < 6; col++)
            {
                //worksheet.Cells[1, col + 1].Value = filteredData.Columns[col].ColumnName;
                worksheet.Cells[1, 1].Value = "MHz";
                worksheet.Cells[1, 2].Value = "";
                worksheet.Cells[1, 3].Value = "S11 - dB";
                worksheet.Cells[1, 4].Value = "S21 - dB";
                worksheet.Cells[1, 5].Value = "S12 - dB";
                worksheet.Cells[1, 6].Value = "S22 - dB";
            }
            // Verileri ekle
            for (int row = 0; row < filteredData.Rows.Count; row++)
            {
                for (int col = 0; col < filteredData.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = Convert.ToDouble(filteredData.Rows[row][col]);
                }
            }
            worksheet.Column(2).Hidden = true;
        }

        private void SaveGrafik(ExcelPackage package, Chart chart, string sheetName)
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            int rowIndex = 1; // Başlangıç satırı
            int colIndex = 10; // Başlangıç sütunu

            // Grafikleri belirtilen hücreye eklemek için resim olarak ekleyin
            using (MemoryStream ms = new MemoryStream())
            {
                chart.SaveImage(ms, ChartImageFormat.Png); // Grafik resmini PNG olarak kaydedin
                var picture = worksheet.Drawings.AddPicture("GrafikResmi", ms);
                picture.SetPosition(rowIndex, 0, colIndex, 0); // Grafik resminin konumunu ayarlayın
                picture.SetSize(400, 300); // Grafik resminin boyutunu ayarlayın
            }
        }

        private void VerileriGrafikteGoster(DataTable dataTable)
        {
            chart1.Series.Clear();

            chart1.ChartAreas[0].AxisX.Title = "MHz";
            chart1.ChartAreas[0].AxisX.Minimum = int.Parse(comboBoxAralikMin.Text);
            chart1.ChartAreas[0].AxisX.Maximum = int.Parse(comboBoxAralikMax.Text);

            chart1.ChartAreas[0].AxisY.Maximum = 20;
            chart1.ChartAreas[0].AxisY.Minimum = -40;

            chart1.ChartAreas[0].AxisY.Title = "dB";

            string[] headerText = new string[dataGridView1.Columns.Count];

            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                headerText[i] = dataGridView1.Columns[i].HeaderText;
            }
            string[] denemeColumnName = { "Column1", "Column2", "Column3", "Column4", "Column5", "Column6" };
            string[] seriesNames = { "S11 - dB", "S21 - dB", "S12 - dB", "S22 - dB" };

            Series series1 = new Series("S11 - dB");
            Series series2 = new Series("S21 - dB");
            Series series3 = new Series("S12 - dB");
            Series series4 = new Series("S22 - dB");

            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series1);
            series1.ChartType = SeriesChartType.Line; // Çizgi grafiği
            series1.BorderWidth = 4; // Çizgi kalınlığı
            series2.ChartType = SeriesChartType.Line; // Çizgi grafiği
            series2.BorderWidth = 4; // Çizgi kalınlığı
            series3.ChartType = SeriesChartType.Line; // Çizgi grafiği
            series3.BorderWidth = 4; // Çizgi kalınlığı
            series4.ChartType = SeriesChartType.Line; // Çizgi grafiği
            series4.BorderWidth = 4; // Çizgi kalınlığı

            
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                double xDegeri = Convert.ToDouble(dataTable.Rows[i]["Column1"]);

                for (int j = 0; j < 4; j++)
                {
                    double yDegeri = 0; // Varsayılan değer

                    if (dataTable.Rows[i][denemeColumnName[j + 2]] != DBNull.Value)
                    {
                        if (double.TryParse(dataTable.Rows[i][denemeColumnName[j + 2]].ToString(), out double parsedValue))
                        {
                            yDegeri = parsedValue;
                        }
                        else
                        {
                            // Uygun bir sayısal değer elde edilemedi, hata durumu veya varsayılan işlemler burada yapılabilir.
                        }
                    }

                    switch (j)
                    {
                        case 0:
                            series1.Points.AddXY(xDegeri, yDegeri);
                            break;
                        case 1:
                            series2.Points.AddXY(xDegeri, yDegeri);
                            break;
                        case 2:
                            series3.Points.AddXY(xDegeri, yDegeri);
                            break;
                        case 3:
                            series4.Points.AddXY(xDegeri, yDegeri);
                            break;
                        default:
                            break;
                    }
                }
            }


        }
    }
}
