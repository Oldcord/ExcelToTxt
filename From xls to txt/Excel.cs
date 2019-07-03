using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel=Microsoft.Office.Interop.Excel;

namespace From_xls_to_txt {
    public partial class Excel : Form {

        _Excel.Application  excel = new _Excel.Application();
        

        Workbook            excelWorkbook;
        Worksheet           excelWorksheet;
        Sheets              excelSheets;
        Range               Range_Number;

        string              pathFile,       //путь файла
                            text_NameFile,  //имя файла
                            text_excel;     //текст файла

        public Excel() { InitializeComponent(); }

        private void BtnOpen_Click(object sender, EventArgs e) {
            //Открыть файл, записать путь файла.
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Open file(*.xls)|*.xls|Open file(*.xlsx)|*.xlsx", ValidateNames = true } ) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    pathFile = Path.GetFullPath(ofd.FileName);
                    textPathFile.Text = Path.GetFullPath(ofd.FileName);
                    cboSheet.Items.Clear();
                    excelWorkbook = excel.Workbooks.Open(pathFile);
                    excelSheets = excelWorkbook.Worksheets;
                    for (int i = 1; i <= excelSheets.Count; i++) {
                        cboSheet.Items.Add(excelSheets[i].Name.ToString());
                    }
                    excelWorkbook.Close(false);
                }
            }
        }

        //Сохранить текстовый фаил *.txt.
        private void ButtonSave_Click(object sender, EventArgs e) {
            SaveFileDialog sfd =    new SaveFileDialog();
            sfd.FileName =          text_NameFile;
            sfd.Title =             "Save file (*.txt)";
            sfd.Filter =            "Save file (*.txt)|*.txt";

            if (sfd.ShowDialog() == DialogResult.OK) {
                //textReadFile.SaveFile(sfd.FileName, RichTextBoxStreamType.PlainText);
                File.WriteAllText(sfd.FileName, text_excel);
            }
        }

        //Выбрать активный лист Excel книги.Открыть страницу книги.
        private void CboSheet_SelectedIndexChanged(object sender, EventArgs e) {
            int n_Columns = 0;  //количество используемых колонок
            int n_Rows = 0;     //количество используемых строк

            try {
                excelWorkbook =     excel.Workbooks.Open(pathFile);
                excelSheets =       excelWorkbook.Worksheets;
                excelWorksheet =    excelSheets.Item[ cboSheet.SelectedIndex + 1 ];
                Range_Number =      excelWorksheet.UsedRange;
                text_NameFile =     excelWorkbook.Name;
                n_Columns =         Range_Number.Columns.Count;
                n_Rows =            Range_Number.Rows.Count;

                //Проверка ячеек на данные
                ReadeExcel(n_Columns, n_Rows);
            }
            catch (Exception) {
                excelWorkbook.Close(false);
                excel.Quit();
                //throw;
            }
        }

        //Проверка ячеек на данные, если есть данные записать в строку text_excel.
        private void ReadeExcel(int column, int row) {
            string text_cells = "";
            text_excel = "";
            for (int x = 1; x <= row; x++) {
                for (int y = 1; y <= column; y++) {
                    if ((excelWorksheet.Cells[x, y].Value2 ?? string.Empty).ToString() != "") {
                        text_cells += excelWorksheet.Cells[x, y].Value2.ToString() + "\t";
                    }
                }
                if (text_cells != "") {
                    text_excel += text_cells + "\n";
                }
                text_cells = "";
            }
            textReadFile.Text = text_excel;
            excelWorkbook.Close();
            excel.Quit();
        }
    }
}
