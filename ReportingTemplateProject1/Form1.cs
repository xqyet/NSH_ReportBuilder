using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CsvHelper;
using ClosedXML.Excel;
using System.Globalization;
using System.Collections.Generic;
using System.Drawing;

namespace ReportingTemplateProject1
{
    public partial class Form1 : Form
    {
        private DataTable csvData;
        private DataTable ordProcedureCptReference;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxReports.Items.Clear(); // Keep the dropdown empty on load

            comboBoxReports.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxReports.DrawItem += comboBoxReports_DrawItem;

            // Attach the DropDown event to populate items when clicked
            comboBoxReports.DropDown += comboBoxReports_DropDown;

            LoadOrdProcedureCptReference();
        }

        // Event handler to populate the dropdown when it's clicked
        private void comboBoxReports_DropDown(object sender, EventArgs e)
        {
            // Only populate if the dropdown is currently empty
            if (comboBoxReports.Items.Count == 0)
            {
                comboBoxReports.Items.Add("PAS Surgery Schedule");
                comboBoxReports.Items.Add("Birth Log Report"); // Add the new option here
            }
        }

        private void comboBoxReports_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            ComboBox comboBox = sender as ComboBox;
            string text = comboBox.Items[e.Index].ToString();

            // Set default background and text color
            e.DrawBackground();

            Color textColor = e.ForeColor;

            // Set custom color for selected item
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                textColor = Color.Black; // Set selected item text color to black
                e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds); // Highlight the background
            }
            else
            {
                e.Graphics.FillRectangle(new SolidBrush(e.BackColor), e.Bounds); // Normal background color
            }

            // Center the text vertically and horizontally in the ComboBox item
            var textSize = e.Graphics.MeasureString(text, comboBox.Font);
            var location = new System.Drawing.Point(
                e.Bounds.X + (e.Bounds.Width - (int)textSize.Width) / 2,
                e.Bounds.Y + (e.Bounds.Height - (int)textSize.Height) / 2
            );

            // Draw the text with the appropriate color
            using (var textBrush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(text, comboBox.Font, textBrush, location);
            }

            e.DrawFocusRectangle();
        }

        private void LoadOrdProcedureCptReference()
        {
            try
            {
                using (var stream = File.Open(@"P:\Patient_Access_All\Cherokee\Schedules\Room Control\DAILY SCHEDULE\CPT_MAPPING.xlsx", FileMode.Open, FileAccess.Read))
                using (var excelReader = new XLWorkbook(stream))
                {
                    var worksheet = excelReader.Worksheet(1);
                    ordProcedureCptReference = new DataTable();

                    ordProcedureCptReference.Columns.Add("ORD_PROCEDURE", typeof(string));
                    ordProcedureCptReference.Columns.Add("CPT", typeof(string));

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var dataRow = ordProcedureCptReference.NewRow();
                        dataRow["ORD_PROCEDURE"] = row.Cell(1).GetString();
                        dataRow["CPT"] = string.IsNullOrWhiteSpace(row.Cell(2).GetString()) ? string.Empty : row.Cell(2).GetString();
                        ordProcedureCptReference.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load ORD_PROCEDURE to CPT reference file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnImportCsv_Click(object sender, EventArgs e)
        {
            // Check if a report is selected before proceeding
            if (comboBoxReports.SelectedItem == null)
            {
                MessageBox.Show("Please select a report you'd like to generate from the dropdown.", "No Report Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Proceed with the CSV or Excel import logic if a report is selected
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                    string fileExtension = Path.GetExtension(openFileDialog.FileName).ToLower();

                    // Handle CSV files
                    if (fileExtension == ".csv")
                    {
                        LoadCsvData(openFileDialog.FileName);
                        lblStatus.Text = "CSV loaded successfully.";
                    }
                    // Handle Excel files
                    else if (fileExtension == ".xlsx")
                    {
                        LoadExcelDataForCptMapping(openFileDialog.FileName);
                        lblStatus.Text = "Excel file loaded successfully, CPT mapping applied where necessary.";
                    }
                    else
                    {
                        MessageBox.Show("Please select a valid CSV or Excel file.", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        // Method to load Excel file and apply CPT mapping only if PAS Surgery Schedule is selected
        private void LoadExcelDataForCptMapping(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (var excelReader = new XLWorkbook(stream))
                {
                    var worksheet = excelReader.Worksheet(1); // Assuming data is on the first worksheet
                    var dataTable = new DataTable();

                    // Load the Excel column names into the DataTable
                    foreach (var cell in worksheet.Row(1).CellsUsed())
                    {
                        dataTable.Columns.Add(cell.GetString());
                    }

                    // Load the Excel data into the DataTable
                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var dataRow = dataTable.NewRow();
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            dataRow[i] = row.Cell(i + 1).GetString();
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    csvData = dataTable; // Assign to the global csvData variable

                    // Only map CPT codes if PAS Surgery Schedule is selected
                    if (comboBoxReports.SelectedItem?.ToString() == "PAS Surgery Schedule")
                    {
                        MapCptCodes(); // Apply CPT mapping logic
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load Excel file and map CPT codes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadCsvData(string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                using (var dr = new CsvDataReader(csv))
                {
                    csvData = new DataTable();
                    csvData.Load(dr);
                }
            }

            // Only map CPT codes if PAS Surgery Schedule is selected
            if (comboBoxReports.SelectedItem?.ToString() == "PAS Surgery Schedule")
            {
                MapCptCodes(); // Apply CPT mapping logic
            }
        }

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (csvData == null)
                {
                    MessageBox.Show("Please import a CSV or Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var selectedReport = comboBoxReports.SelectedItem?.ToString();

                if (selectedReport == null)
                {
                    MessageBox.Show("Please select a report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Only apply CPT-related logic for PAS Surgery Schedule
                if (selectedReport == "PAS Surgery Schedule")
                {
                    // Check if CPT column exists for PAS Surgery Schedule
                    if (!csvData.Columns.Contains("CPT"))
                    {
                        MessageBox.Show("CPT column is missing in the data. Please verify the CSV file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    GeneratePasSurgeryScheduleReport(); // Generate PAS Surgery Schedule
                }
                else if (selectedReport == "Birth Log Report")
                {
                    // No CPT logic for Birth Log Report
                    GenerateBirthLogReport(); // Generate Birth Log Report
                }
                else
                {
                    MessageBox.Show("Selected report logic not yet implemented.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                comboBoxReports.SelectedIndex = -1; // Clear the selection after generating the report
                comboBoxReports.ResetText();        // Ensure the displayed text is cleared
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GeneratePasSurgeryScheduleReport()
        {
            try
            {
                var columnsToInclude = new string[]
                {
                    "PATIENT_NAME", "DOB", "FIN", "PT TYPE", "DEPT",
                    "ORD_PROCEDURE", "CPT", "REASON_FOR_VISIT",
                    "ATTENDING_PHYS", "APPT_DT_TM"
                };

                // Dynamically add PT TYPE and DEPT columns if they do not exist
                if (!csvData.Columns.Contains("PT TYPE"))
                {
                    csvData.Columns.Add("PT TYPE", typeof(string));
                }

                if (!csvData.Columns.Contains("DEPT"))
                {
                    csvData.Columns.Add("DEPT", typeof(string));
                }

                // If the file was a CSV file (which contains the LOCATION column), populate DEPT based on LOCATION
                if (csvData.Columns.Contains("LOCATION"))
                {
                    foreach (DataRow row in csvData.Rows)
                    {
                        string location = row["LOCATION"]?.ToString() ?? string.Empty;
                        switch (location)
                        {
                            case "CARDIOLOGY-NSC":
                                row["DEPT"] = "CVC";
                                break;
                            case "CATHLAB ARU-NSC":
                            case "CATH LAB-NSC":
                                row["DEPT"] = "CCIL";
                                break;
                            case "NEURODIAG-NSC":
                                row["DEPT"] = "EEGC";
                                break;
                            case "NSC Endoscopy Cherokee":
                                row["DEPT"] = "GIC";
                                break;
                            case "NSC Operating Room":
                                row["DEPT"] = "ORC";
                                break;
                            case "RAD DX IMG-NSC":
                                row["DEPT"] = "RADC";
                                break;
                            case "PULMONARY-NSC":
                            case "PULM DIAG-NSC":
                                row["DEPT"] = "RESP";
                                break;
                            default:
                                row["DEPT"] = location; // If no match, keep the location as dept
                                break;
                        }

                        // Only populate PT TYPE if there is a valid value, no "Default" value
                        if (row["PT TYPE"] == DBNull.Value || string.IsNullOrWhiteSpace(row["PT TYPE"].ToString()))
                        {
                            row["PT TYPE"] = string.Empty;
                        }
                    }
                }

                // Ensure relevant columns exist
                if (!csvData.Columns.Contains("CPT"))
                {
                    MessageBox.Show("CPT column is missing in the data. Please verify the CSV file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Populate missing CPT codes using the ORD_PROCEDURE to CPT mapping
                MapCptCodes();

                var filteredData = csvData.DefaultView.ToTable(false, columnsToInclude);

                // Sort the data alphabetically by PATIENT_NAME
                DataView dv = filteredData.DefaultView;
                dv.Sort = "PATIENT_NAME ASC";
                filteredData = dv.ToTable();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        InsertDataIntoTemplate(filteredData, saveFileDialog.FileName);
                        lblStatus.Text = "PAS Surgery Schedule report generated successfully.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate PAS Surgery Schedule report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Method to map CPT codes based on ORD_PROCEDURE
        private void MapCptCodes()
        {
            // Ensure CPT column is writable
            if (csvData.Columns.Contains("CPT"))
            {
                csvData.Columns["CPT"].ReadOnly = false;
            }

            // Map missing CPT codes for CSV and Excel data
            foreach (DataRow row in csvData.Rows)
            {
                if (string.IsNullOrWhiteSpace(row["CPT"].ToString()))
                {
                    var ordProcedure = row["ORD_PROCEDURE"].ToString();
                    var cptMapping = ordProcedureCptReference.AsEnumerable()
                        .FirstOrDefault(r => r.Field<string>("ORD_PROCEDURE") == ordProcedure);

                    row["CPT"] = cptMapping?["CPT"]?.ToString() ?? string.Empty;
                }
            }
        }

        // Insert data into Excel template for PAS Surgery Schedule
        private void InsertDataIntoTemplate(DataTable dataTable, string filePath)
        {
            using (var templateWorkbook = new XLWorkbook(@"P:\Patient_Access_All\Cherokee\Schedules\Room Control\FORMS\SCHEDULE_SHELL.xlsx"))
            {
                var worksheet = templateWorkbook.Worksheet(1); // Assuming data goes into the first sheet

                // Define the starting row for data insertion (adjust according to your template)
                int startRow = 2; // Example: Start inserting data at row 2

                // Variable to keep track of the previous patient name
                string previousPatientName = string.Empty;

                // Insert data from the filtered DataTable into the template
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    string currentPatientName = dataTable.Rows[i]["PATIENT_NAME"].ToString();

                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        var cellValue = dataTable.Rows[i][j]?.ToString();

                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            var cell = worksheet.Cell(startRow + i, j + 1);
                            cell.Value = cellValue;

                            // Set the row height to 24
                            worksheet.Row(startRow + i).Height = 24;

                            // Enable text wrapping if cell content exceeds a certain length
                            if (j == 0 && cell.GetText().Length > 20) // Assuming PATIENT_NAME is the first column (index 0)
                            {
                                cell.Style.Alignment.WrapText = true;
                                worksheet.Row(startRow + i).Height = 32;
                            }

                            // Apply bold formatting for new patient names
                            if (j == 0 && currentPatientName != previousPatientName)
                            {
                                cell.Style.Font.Bold = true;
                            }

                            // Center and bold the PT TYPE column
                            if (j == 3) // Assuming PT TYPE is the fourth column (index 3)
                            {
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                cell.Style.Font.Bold = true;
                            }
                        }
                    }

                    // Update the previous patient name
                    previousPatientName = currentPatientName;
                }

                // Adjust column widths and wrap text in ORD_PROCEDURE column to prevent overflow into CPT column
                worksheet.Column(6).Width = 25;
                worksheet.Column(6).Style.Alignment.WrapText = true;

                // Center the DEPT column
                worksheet.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Set custom margins and orientation
                worksheet.PageSetup.Margins.Top = 0.54;
                worksheet.PageSetup.Margins.Bottom = 0.2;
                worksheet.PageSetup.Margins.Right = 0.010416666667;
                worksheet.PageSetup.Margins.Footer = 0.1;
                worksheet.PageSetup.CenterHorizontally = true;
                worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                // Set report date in the header
                var reportDate = DateTime.Now.AddDays(1); // Next day's date
                worksheet.PageSetup.Header.Center.AddText($"{reportDate:MMMM dd, yyyy}".ToUpper() + $"\n{reportDate:dddd}".ToUpper());

                // Set current date and time in the right section of the header
                var currentTime = DateTime.Now;
                worksheet.PageSetup.Header.Right.AddText($"RAN {currentTime:MM/dd/yyyy} @{currentTime:HHmm}");

                // Add page numbers at the bottom center
                worksheet.PageSetup.Footer.Center.AddText("Page &P of &N");

                // Save the final report
                templateWorkbook.SaveAs(filePath);
            }
        }

        private void GenerateBirthLogReport()
        {
            try
            {
                // Step 1: Check if CSV file is loaded
                if (csvData == null)
                {
                    MessageBox.Show("Please import a CSV file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Step 2: Path for the Birth Log Book Report Shell
                string excelTemplatePath = @"P:\Patient_Access_All\Cherokee\Schedules\Room Control\CAROLINE\BIRTH LOG BOOK REPORT SHELL.xlsx";

                // Load the Excel template
                using (var stream = File.Open(excelTemplatePath, FileMode.Open, FileAccess.Read))
                using (var excelReader = new XLWorkbook(stream))
                {
                    var worksheet = excelReader.Worksheet(1); // Assuming the data goes into the first sheet

                    // Step 3: Map columns from CSV to the required columns in the report
                    var columnMapping = new Dictionary<string, string>()
                    {
                        {"Name(Mother)", "MOTHER'S NAME"},
                        {"Admit_Date/Time", "MOTHER'S ADMIT DATE"},
                        {"Age", "AGE"},
                        {"EGA", "EGA"},
                        {"Birth_Date/Time", "BABY'S BIRTH DATE & TIME"},
                        {"Delivery_Type", "DELIVERY TYPE"},
                        {"Sex_of_Baby", "BABY'S SEX"},
                        {"Race", "RACE"}
                    };

                    // Ensure all required columns exist for the Birth Log report
                    foreach (var col in columnMapping.Keys)
                    {
                        if (!csvData.Columns.Contains(col))
                        {
                            MessageBox.Show($"Column {col} is missing in the data. Please verify the CSV file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    // Step 4: Set column widths
                    worksheet.Column(4).Width = 88 / 7.0; // EGA column set to 88 pixels wide
                    worksheet.Column(5).Width = 119 / 7.0; // BABY'S BIRTH DATE & TIME column set to 119 pixels wide

                    // Define the starting row for data insertion
                    int startRow = 2; // Assuming row 1 is the header in the Excel template

                    // Step 5: Populate the Excel template with data from the CSV
                    for (int i = 0; i < csvData.Rows.Count; i++)
                    {
                        var row = csvData.Rows[i];

                        // Insert data into the corresponding columns
                        worksheet.Cell(startRow + i, 1).Value = row["Name(Mother)"]?.ToString();
                        worksheet.Cell(startRow + i, 2).Value = row["Admit_Date/Time"]?.ToString();
                        worksheet.Cell(startRow + i, 3).Value = row["Age"]?.ToString();
                        worksheet.Cell(startRow + i, 4).Value = row["EGA"]?.ToString();
                        worksheet.Cell(startRow + i, 5).Value = row["Birth_Date/Time"]?.ToString();
                        worksheet.Cell(startRow + i, 6).Value = row["Delivery_Type"]?.ToString();
                        worksheet.Cell(startRow + i, 7).Value = row["Sex_of_Baby"]?.ToString();
                        worksheet.Cell(startRow + i, 8).Value = row["Race"]?.ToString();

                        // Set MOTHER'S NAME column (column 1) to Arial and left-aligned
                        worksheet.Cell(startRow + i, 1).Style.Font.FontName = "Arial";
                        worksheet.Cell(startRow + i, 1).Style.Font.Bold = false; // Ensure it's not bold
                        worksheet.Cell(startRow + i, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                        // Set other columns to Calibri, size 11, and center-aligned (except for MOTHER'S NAME column)
                        for (int j = 2; j <= 8; j++)
                        {
                            worksheet.Cell(startRow + i, j).Style.Font.FontName = "Calibri";
                            worksheet.Cell(startRow + i, j).Style.Font.FontSize = 11;
                            worksheet.Cell(startRow + i, j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        // Specific formatting for the EGA column (column 4)
                        worksheet.Cell(startRow + i, 4).Style.Font.FontSize = 12; // Font size 12
                        worksheet.Cell(startRow + i, 4).Style.Font.Bold = false; // Ensure it's not bold

                        // Specific formatting for the AGE column (column 3)
                        worksheet.Cell(startRow + i, 3).Style.Font.FontSize = 12; // Font size 12
                    }

                    // Get current date and time
                    var currentTime = DateTime.Now;
                    string ranTime = $"RAN {currentTime:MM/dd/yyyy} @{currentTime:HHmm}";

                    // Set the RAN time in the header of the worksheet
                    worksheet.PageSetup.Header.Right.AddText(ranTime);

                    // Step 6: Save the populated report
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            excelReader.SaveAs(saveFileDialog.FileName);
                            MessageBox.Show("Birth Log Report generated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate the Birth Log Report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}


