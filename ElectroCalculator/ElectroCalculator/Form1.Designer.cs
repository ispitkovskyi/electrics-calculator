using System;
using System.IO;
using System.Windows.Forms;
using myExcel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ElectroCalculator
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TextBox fldMonthNumber;
        private System.Windows.Forms.Label lblMonthNumber;
        private System.Windows.Forms.Button btnOpenFile;
        private FileStream workbookStream;
        private string workbookFilePath;
        private int startRowIndex;


        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.fldMonthNumber = new System.Windows.Forms.TextBox();
            this.lblMonthNumber = new System.Windows.Forms.Label();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.fldColumnTitle = new System.Windows.Forms.TextBox();
            this.fldDateColumnIndex = new System.Windows.Forms.TextBox();
            this.lblDateColumn = new System.Windows.Forms.Label();
            this.lblDataColumnTitle = new System.Windows.Forms.Label();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.lblRowNum = new System.Windows.Forms.Label();
            this.lblColNum = new System.Windows.Forms.Label();
            this.lblStartDateRow = new System.Windows.Forms.Label();
            this.lblEndDateRow = new System.Windows.Forms.Label();
            this.lblFirstDataRow = new System.Windows.Forms.Label();
            this.fldFirstDataRow = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "OpenFile";
            this.openFileDialog.InitialDirectory = "%userprofile%\\Desktop";
            // 
            // fldMonthNumber
            // 
            this.fldMonthNumber.AccessibleName = "monthFld";
            this.fldMonthNumber.Location = new System.Drawing.Point(351, 46);
            this.fldMonthNumber.Name = "fldMonthNumber";
            this.fldMonthNumber.Size = new System.Drawing.Size(100, 20);
            this.fldMonthNumber.TabIndex = 0;
            this.fldMonthNumber.Text = "9";
            // 
            // lblMonthNumber
            // 
            this.lblMonthNumber.AutoSize = true;
            this.lblMonthNumber.Location = new System.Drawing.Point(19, 49);
            this.lblMonthNumber.Name = "lblMonthNumber";
            this.lblMonthNumber.Size = new System.Drawing.Size(318, 13);
            this.lblMonthNumber.TabIndex = 1;
            this.lblMonthNumber.Text = "Номер месяца для расчета (1, 2 ... 12). Выбрать только один:";
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(22, 168);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(75, 23);
            this.btnOpenFile.TabIndex = 2;
            this.btnOpenFile.Text = "Open File...";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // fldColumnTitle
            // 
            this.fldColumnTitle.Location = new System.Drawing.Point(379, 124);
            this.fldColumnTitle.Name = "fldColumnTitle";
            this.fldColumnTitle.Size = new System.Drawing.Size(129, 20);
            this.fldColumnTitle.TabIndex = 3;
            this.fldColumnTitle.Text = "+kWh Sys +Wh";
            // 
            // fldDateColumnIndex
            // 
            this.fldDateColumnIndex.Location = new System.Drawing.Point(351, 20);
            this.fldDateColumnIndex.Name = "fldDateColumnIndex";
            this.fldDateColumnIndex.Size = new System.Drawing.Size(100, 20);
            this.fldDateColumnIndex.TabIndex = 4;
            this.fldDateColumnIndex.Text = "A";
            // 
            // lblDateColumn
            // 
            this.lblDateColumn.AutoSize = true;
            this.lblDateColumn.Location = new System.Drawing.Point(19, 23);
            this.lblDateColumn.Name = "lblDateColumn";
            this.lblDateColumn.Size = new System.Drawing.Size(314, 13);
            this.lblDateColumn.TabIndex = 5;
            this.lblDateColumn.Text = "Название колонки содержащей даты (Пример: А, В или АZ):";
            // 
            // lblDataColumnTitle
            // 
            this.lblDataColumnTitle.AutoSize = true;
            this.lblDataColumnTitle.Location = new System.Drawing.Point(19, 127);
            this.lblDataColumnTitle.Name = "lblDataColumnTitle";
            this.lblDataColumnTitle.Size = new System.Drawing.Size(354, 13);
            this.lblDataColumnTitle.TabIndex = 6;
            this.lblDataColumnTitle.Text = "Текст заглавия колонки(ок) для которых нужно вычислить разницу:";
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(113, 173);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(77, 13);
            this.lblFilePath.TabIndex = 7;
            this.lblFilePath.Text = "Путь к файлу:";
            // 
            // lblRowNum
            // 
            this.lblRowNum.AutoSize = true;
            this.lblRowNum.Location = new System.Drawing.Point(22, 208);
            this.lblRowNum.Name = "lblRowNum";
            this.lblRowNum.Size = new System.Drawing.Size(90, 13);
            this.lblRowNum.TabIndex = 8;
            this.lblRowNum.Text = "RowsProcessed: ";
            // 
            // lblColNum
            // 
            this.lblColNum.AutoSize = true;
            this.lblColNum.Location = new System.Drawing.Point(22, 238);
            this.lblColNum.Name = "lblColNum";
            this.lblColNum.Size = new System.Drawing.Size(45, 13);
            this.lblColNum.TabIndex = 9;
            this.lblColNum.Text = "Column:";
            // 
            // lblStartDateRow
            // 
            this.lblStartDateRow.AutoSize = true;
            this.lblStartDateRow.Location = new System.Drawing.Point(180, 208);
            this.lblStartDateRow.Name = "lblStartDateRow";
            this.lblStartDateRow.Size = new System.Drawing.Size(80, 13);
            this.lblStartDateRow.TabIndex = 10;
            this.lblStartDateRow.Text = "StartDateRow: ";
            // 
            // lblEndDateRow
            // 
            this.lblEndDateRow.AutoSize = true;
            this.lblEndDateRow.Location = new System.Drawing.Point(338, 208);
            this.lblEndDateRow.Name = "lblEndDateRow";
            this.lblEndDateRow.Size = new System.Drawing.Size(74, 13);
            this.lblEndDateRow.TabIndex = 11;
            this.lblEndDateRow.Text = "EndDateRow:";
            // 
            // lblFirstDataRow
            // 
            this.lblFirstDataRow.AutoSize = true;
            this.lblFirstDataRow.Location = new System.Drawing.Point(19, 78);
            this.lblFirstDataRow.Name = "lblFirstDataRow";
            this.lblFirstDataRow.Size = new System.Drawing.Size(262, 13);
            this.lblFirstDataRow.TabIndex = 12;
            this.lblFirstDataRow.Text = "Номер строки где начинаются измерения (числа):";
            // 
            // fldFirstDataRow
            // 
            this.fldFirstDataRow.Location = new System.Drawing.Point(351, 75);
            this.fldFirstDataRow.Name = "fldFirstDataRow";
            this.fldFirstDataRow.Size = new System.Drawing.Size(100, 20);
            this.fldFirstDataRow.TabIndex = 13;
            this.fldFirstDataRow.Text = "6";
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(520, 270);
            this.Controls.Add(this.fldFirstDataRow);
            this.Controls.Add(this.lblFirstDataRow);
            this.Controls.Add(this.lblEndDateRow);
            this.Controls.Add(this.lblStartDateRow);
            this.Controls.Add(this.lblColNum);
            this.Controls.Add(this.lblRowNum);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.lblDataColumnTitle);
            this.Controls.Add(this.lblDateColumn);
            this.Controls.Add(this.fldDateColumnIndex);
            this.Controls.Add(this.fldColumnTitle);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.lblMonthNumber);
            this.Controls.Add(this.fldMonthNumber);
            this.Name = "Form1";
            this.Text = "ElectricCalculator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog.Filter = "Text Files (.xls)|*.xls|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

            openFileDialog.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                // Open the selected file to read.
                try
                {
                    workbookFilePath = openFileDialog.FileName;
                    lblFilePath.Text = workbookFilePath;
                    workbookStream = new FileStream(workbookFilePath, FileMode.Open);
                }
                catch (IOException exception)
                {
                    MessageBox.Show("Excel file is currently opened by another application. Close file before running this application", "Error", MessageBoxButtons.OK);
                }

                using (StreamReader reader = new StreamReader(workbookStream))
                {
                    // Read the first line from the file and write it the textbox.
                    initCaclulationParameters();
                    parseExcel(workbookFilePath);
                }
                workbookStream.Close();
            }
        }

        int month;
        string datesColumnIndex;
        string dataColumnsTitle;
        private void initCaclulationParameters()
        {
            month = Int32.Parse(this.fldMonthNumber.Text);
            datesColumnIndex = this.fldDateColumnIndex.Text;
            dataColumnsTitle = this.fldColumnTitle.Text;
            startRowIndex = Int32.Parse(this.fldFirstDataRow.Text);
        }

        private void parseExcel(string filePath) {
            myExcel.Application excel = new myExcel.Application();
            myExcel.Workbook wbook;
            myExcel.Worksheet wsheet;

            wbook = excel.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            wsheet = (myExcel.Worksheet)wbook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range datesCol = wsheet.Columns[datesColumnIndex];//wsheet.Columns[1];
            int rangeStart = -1;
            int rangeEnd = -1;

            findDatesRange(datesCol, ref rangeStart, ref rangeEnd);
            msg += "RANGE DEFINED by DATE. START ROW = " + rangeStart + "  END ROW = " + rangeEnd + "\n";

                if (rangeStart < 0 || rangeEnd < 0)
                    throw new Exception("Start or end of diapason were not found");

            Microsoft.Office.Interop.Excel.Range measurementsColumns = wsheet.Columns;
            calculateDifferences(measurementsColumns, rangeStart, rangeEnd);
            excel.Workbooks.Close();
            MessageBox.Show(msg, "Result", MessageBoxButtons.OK);
        }


        private void findDatesRange (myExcel.Range datesCol, ref int rangeStart, ref int rangeEnd) {
            int monthValue;
            Boolean startFound = false;
            for (int i = startRowIndex; ; i++)
            {
                myExcel.Range dcell = datesCol.EntireColumn.Cells.get_Item(i);
                try
                {
                    string[] date_parts = dcell.get_Value().ToString().Split(new String[] { " -" }, StringSplitOptions.RemoveEmptyEntries);
                    string date = date_parts[0] + date_parts[1];
                    DateTime dateTime = DateTime.Parse(date, null); //DateTime.ParseExact(date, "dd MMM yyyy HH:mm:ss", null);
                    monthValue = dateTime.Month;
                    if (monthValue == month && !startFound)
                    {
                        rangeStart = i;
                        startFound = true;
                        lblStartDateRow.Text = "StartDateRow: " + dcell.Address.Split(':')[0];
                        System.Console.Out.WriteLine("Start row: " + i);
                    }
                    if (monthValue != month && startFound) //next month range starts
                    {
                        rangeEnd = i - 1;
                        lblEndDateRow.Text = "EndDateRow: " + datesCol.EntireColumn.Cells.get_Item(i - 1).Address.Split(':')[0];
                        System.Console.Out.WriteLine("End row: " + (i - 1));
                        break;
                    }
                    lblRowNum.Text = "Rows Processed: " + i;
                }
                catch (Exception e)
                { //only one month in diapason, so blank cell (w/o no date) occurs
                    if (!startFound)
                        MessageBox.Show("Could not identify range for the specified month: " + month, "ERROR", MessageBoxButtons.OK);

                    System.Console.Out.WriteLine("Set range end here, because could not get date from cell");
                    rangeEnd = i - 1;
                    lblEndDateRow.Text = "EndDateRow: " + datesCol.EntireColumn.Cells.get_Item(i - 1).Address.Split(':')[0];
                    System.Console.Out.WriteLine("End row: " + (i - 1));
                    break;
                }
            }       
        }

        private void calculateDifferences(myExcel.Range columns, int rangeStart, int rangeEnd) {
            myExcel.Range dataCell;
            int columnsCounter = 0;
            foreach (myExcel.Range col in columns)
            {
                myExcel.Range cell = col.Find(dataColumnsTitle);
                if (cell != null && cell.Count > 0)
                {
                    lblColNum.Text = "Measurements Column" + col.Address + ": " + columnsCounter++;

                    dataCell = col.EntireColumn.Cells.get_Item(rangeStart);
                    object startValueStr = dataCell.get_Value();
                    int startValue = Int32.Parse(startValueStr.ToString());

                    dataCell = col.EntireColumn.Cells.get_Item(rangeEnd);
                    object endValueObj = (object)dataCell.get_Value();
                    int endValue = Int32.Parse(endValueObj.ToString());

                    appendToMessage(col.Address, startValue, endValue);
                }
            }
        }



        private static string msg = "";
        private void appendToMessage(string columnAddr, int startValue, int endValue) {
            msg = msg + "Column: " + columnAddr.Split(':')[0] + "\n\t Start value: " + startValue + " --> End value: " + endValue + "\n\t\t\t\t\t Result: " + (endValue - startValue) + "\n";            
        }

        private TextBox fldColumnTitle;
        private TextBox fldDateColumnIndex;
        private Label lblDateColumn;
        private Label lblDataColumnTitle;
        private Label lblFilePath;
        private Label lblRowNum;
        private Label lblColNum;
        private Label lblStartDateRow;
        private Label lblEndDateRow;
        private Label lblFirstDataRow;
        private TextBox fldFirstDataRow;


    }
}

