using OfficeOpenXml;
using System.Data;

namespace ExcelReconApp;

public partial class Form1 : Form
{
    private DataTable worksheet1Data = new DataTable();
    private DataTable worksheet2Data = new DataTable();
    private DataTable reconResultsData = new DataTable();

    public Form1()
    {
        InitializeComponent();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        InitializeReconResultsTable();
    }

    private void InitializeReconResultsTable()
    {
        reconResultsData.Columns.Add("ID", typeof(string));
        reconResultsData.Columns.Add("Sheet1 Total", typeof(double));
        reconResultsData.Columns.Add("Sheet2 Total", typeof(double));
        reconResultsData.Columns.Add("IsMatch", typeof(bool));
        dataGridViewReconResults.DataSource = reconResultsData;
    }

    private void btnLoadWorksheet1_Click(object sender, EventArgs e)
    {
        LoadWorksheet(ref worksheet1Data, dataGridViewWorksheet1, comboBoxIdColumn1, checkedListBoxValueColumns1);
    }

    private void btnLoadWorksheet2_Click(object sender, EventArgs e)
    {
        LoadWorksheet(ref worksheet2Data, dataGridViewWorksheet2, comboBoxIdColumn2, checkedListBoxValueColumns2);
    }

    private void LoadWorksheet(ref DataTable dataTable, DataGridView gridView, ComboBox idColumnComboBox, CheckedListBox valueColumnsListBox)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(openFileDialog.FileName);
                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        
                        dataTable = new DataTable();
                        
                        // Get headers from first row
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            string columnName = worksheet.Cells[1, col].Text;
                            if (string.IsNullOrWhiteSpace(columnName))
                                columnName = $"Column{col}";
                            dataTable.Columns.Add(columnName);
                        }

                        // Load data starting from row 2
                        for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                            {
                                dataRow[col - 1] = worksheet.Cells[row, col].Text;
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        gridView.DataSource = dataTable;
                        
                        // Update column pickers
                        UpdateColumnPickers(dataTable, idColumnComboBox, valueColumnsListBox);
                        
                        MessageBox.Show($"Successfully loaded {dataTable.Rows.Count} rows from {Path.GetFileName(openFileDialog.FileName)}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void UpdateColumnPickers(DataTable dataTable, ComboBox idColumnComboBox, CheckedListBox valueColumnsListBox)
    {
        idColumnComboBox.Items.Clear();
        valueColumnsListBox.Items.Clear();

        foreach (DataColumn column in dataTable.Columns)
        {
            idColumnComboBox.Items.Add(column.ColumnName);
            valueColumnsListBox.Items.Add(column.ColumnName);
        }

        if (idColumnComboBox.Items.Count > 0)
            idColumnComboBox.SelectedIndex = 0;
    }

    private void btnRecon_Click(object sender, EventArgs e)
    {
        // Validate selections
        if (worksheet1Data.Rows.Count == 0 || worksheet2Data.Rows.Count == 0)
        {
            MessageBox.Show("Please load both worksheets before running reconciliation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (comboBoxIdColumn1.SelectedItem == null || comboBoxIdColumn2.SelectedItem == null)
        {
            MessageBox.Show("Please select ID columns for both worksheets.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (checkedListBoxValueColumns1.CheckedItems.Count == 0 || checkedListBoxValueColumns2.CheckedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one value column for both worksheets.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            PerformReconciliation();
            MessageBox.Show($"Reconciliation completed! {reconResultsData.Rows.Count} records processed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error during reconciliation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void PerformReconciliation()
    {
        reconResultsData.Clear();

        string idColumn1 = comboBoxIdColumn1.SelectedItem.ToString()!;
        string idColumn2 = comboBoxIdColumn2.SelectedItem.ToString()!;

        List<string> valueColumns1 = checkedListBoxValueColumns1.CheckedItems.Cast<string>().ToList();
        List<string> valueColumns2 = checkedListBoxValueColumns2.CheckedItems.Cast<string>().ToList();

        // Build dictionaries for quick lookup
        Dictionary<string, double> sheet1Totals = BuildTotalsDictionary(worksheet1Data, idColumn1, valueColumns1);
        Dictionary<string, double> sheet2Totals = BuildTotalsDictionary(worksheet2Data, idColumn2, valueColumns2);

        // Get all unique IDs from both sheets
        HashSet<string> allIds = new HashSet<string>(sheet1Totals.Keys);
        allIds.UnionWith(sheet2Totals.Keys);

        // Process each ID
        foreach (string id in allIds.OrderBy(x => x))
        {
            double sheet1Total = sheet1Totals.ContainsKey(id) ? sheet1Totals[id] : 0;
            double sheet2Total = sheet2Totals.ContainsKey(id) ? sheet2Totals[id] : 0;
            bool isMatch = Math.Abs(sheet1Total - sheet2Total) < 0.01; // Small tolerance for floating point comparison

            DataRow row = reconResultsData.NewRow();
            row["ID"] = id;
            row["Sheet1 Total"] = sheet1Total;
            row["Sheet2 Total"] = sheet2Total;
            row["IsMatch"] = isMatch;
            reconResultsData.Rows.Add(row);
        }
    }

    private Dictionary<string, double> BuildTotalsDictionary(DataTable dataTable, string idColumn, List<string> valueColumns)
    {
        Dictionary<string, double> totals = new Dictionary<string, double>();

        foreach (DataRow row in dataTable.Rows)
        {
            string id = row[idColumn].ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(id))
                continue;

            double total = 0;
            foreach (string valueColumn in valueColumns)
            {
                if (double.TryParse(row[valueColumn].ToString(), out double value))
                {
                    total += value;
                }
            }

            if (totals.ContainsKey(id))
            {
                totals[id] += total;
            }
            else
            {
                totals[id] = total;
            }
        }

        return totals;
    }
}
