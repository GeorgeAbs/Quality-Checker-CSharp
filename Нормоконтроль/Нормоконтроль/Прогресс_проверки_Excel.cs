using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Нормоконтроль
{
    public partial class Прогресс_проверки_Excel : Form
    {
        public Прогресс_проверки_Excel()
        {
            InitializeComponent();
            ruleFile = System.IO.File.ReadAllLines(starting_page_Excel.rulesFile, Encoding.UTF8);
            preparingForChecking.ForeColor = Color.Red;
            checkingIsDone.Visible = false;
            commentsForFinal.Visible = false;
        }
        int numberOfTables;
        int numberOfParagraphOutOfTables = 0;
        int numberOfParagraphInDoc;
        int checkedYesUserRules = 0;
        int currentTableFormZero = 0;
        int currentRow;
        int numberOfRowsInCurrentTable;
        string basicRules = "azaza??";
        string italic = "";
        string textAlighnment = "";
        string indentLeft = "";
        string[] ruleFile;
        bool checkingTables = false, checkingParagraphsOutOfTables = false, checkingUserRules = false;
        bool paraInTempExsistents = false, paraInTempChosed = false;
        Excel.Application appTemp;
        Excel.Workbook tempWorkBook;
        Excel.Application appDoc;
        Excel.Workbook docWorkBook;
        Excel.Worksheet thisSheetDoc;
        Excel.Worksheet thisSheetTemp;

        private void Прогресс_проверки_Excel_Load(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (starting_page_Excel.templateForCheck.Length > 1)
            {
                appTemp = new Microsoft.Office.Interop.Excel.Application();
                tempWorkBook = appTemp.Workbooks.Open(starting_page_Excel.templateForCheck);
            }
            File.Copy(starting_page_Excel.fileForCheck, starting_page_Excel.fileForCheck.Substring(0, starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "\\" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf("\\"), starting_page_Excel.fileForCheck.LastIndexOf(".") - starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "_после проверки" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf(".")));

            appDoc = new Microsoft.Office.Interop.Excel.Application();
            docWorkBook = appDoc.Workbooks.Open(starting_page_Excel.fileForCheck.Substring(0, starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "\\" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf("\\"), starting_page_Excel.fileForCheck.LastIndexOf(".") - starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "_после проверки" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf(".")));


            int i = 0;
            while (i < ruleFile.Length - 1) //подсчет выбранных правил юзерка
            {
                if (ruleFile[i] == "checkState_YES")
                {
                    if (ruleFile[i + 1] != "**??Rule??**")
                    {
                        checkedYesUserRules++;
                    }
                    else
                    {
                        if (ruleFile[i + 1] == "**??Rule??**")
                        {
                            basicRules = basicRules + ruleFile[i - 1] + "??";
                        }
                    }
                }
                i++;
            }
            //проверки
            if (starting_page_Excel.templateForCheck.Length > 1)
            {
                checkFormatingTables(docWorkBook, tempWorkBook);
            }
            checkUserRules(docWorkBook);



            docWorkBook.Close(true, Type.Missing, Type.Missing);
            appDoc.Quit();
            Marshal.ReleaseComObject(docWorkBook);
            Marshal.ReleaseComObject(appDoc);
            if (starting_page_Excel.templateForCheck.Length > 1)
            {
                tempWorkBook.Close();
                appTemp.Quit();
                Marshal.ReleaseComObject(tempWorkBook);
                Marshal.ReleaseComObject(appTemp);
            }

        }
        /*void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (checkingTables == true)
            {
                preparingForChecking.ForeColor = Color.Black;
                checkedTables.ForeColor = Color.Red;
                checkedTablesNumber.Text = currentTableFormZero.ToString() + " из " + numberOfTables.ToString() + " (строка: " + currentRow.ToString() + " из " + numberOfRowsInCurrentTable.ToString() + ")";
            }
            if (checkingParagraphsOutOfTables == true)
            {
                preparingForChecking.ForeColor = Color.Black;
                checkedFormating.ForeColor = Color.Red;
                checkedFormatingNumber.Text = currentRow.ToString() + " из " + numberOfParagraphOutOfTables;
            }
            if (checkingUserRules == true)
            {
                preparingForChecking.ForeColor = Color.Black;
                checkedRules.ForeColor = Color.Red;
                checkUserRulesNumber.Text = currentRow.ToString() + " из " + checkedYesUserRules;
            }
        }*/
        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            checkedRules.ForeColor = Color.Black;
            checkingIsDone.Visible = true;
            commentsForFinal.Visible = true;
        }
        private void checkFormatingTables(Microsoft.Office.Interop.Excel.Workbook docWorkBook, Microsoft.Office.Interop.Excel.Workbook tempWorkBook)
        {
            int sheetsInExcel = docWorkBook.Sheets.Count;
            int ss = 0;
            while (ss < sheetsInExcel)
            {
                ss++;
                var lastCell = docWorkBook.Sheets[ss].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                var lastCellTemp = tempWorkBook.Sheets[ss].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                //разъединение всех ячеек
                Excel.Range unmergeCellsRange = docWorkBook.Sheets[ss].get_Range("A1", NumberToLetters(lastCell.Column) + lastCell.Row.ToString());
                unmergeCellsRange.UnMerge();

                if (lastCell.Row > 0 & lastCell.Column > 0 & lastCellTemp.Row > 0 & lastCellTemp.Column > 0) //если листы не пустые
                {
                    bool sheetsEqual = false;
                    try
                    {
                        thisSheetDoc = docWorkBook.Sheets[ss];
                        thisSheetTemp = tempWorkBook.Sheets[ss];
                        sheetsEqual = true;
                    }
                    catch { }
                    if (sheetsEqual == true) //если есть соответствие листов шаблона документа
                    {
                        int lastRowTemp = lastCellTemp.Row;
                        int checkTitleRow = 0;
                        while (checkTitleRow < lastRowTemp - 1) //проверка шапки (всего строк в шаблоне минус 1 строка примера)
                        {
                            checkTitleRow++;
                            int lastColumnTemp = lastCellTemp.Column;
                            int checkTitleCellColumn = 0;
                            while (checkTitleCellColumn < lastColumnTemp) //проверка строк в шапке
                            {
                                checkTitleCellColumn++;
                                if (thisSheetTemp.Cells[checkTitleRow, checkTitleCellColumn].Text != thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Text)
                                {
                                    thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                    var Text = "Текст не соответствует шаблону;";
                                    thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                }

                            }
                        }
                        checkTitleRow = 0;
                        while (checkTitleRow < lastCell.Row)
                        {
                            checkTitleRow++;
                            int checkTitleCellColumn = 0;
                            while (checkTitleCellColumn < lastCellTemp.Column)
                            {
                                checkTitleCellColumn++;
                                Excel.Range tempCheckNorm;
                                if (checkTitleRow< lastCellTemp.Row)
                                {
                                    tempCheckNorm = thisSheetTemp.Cells[checkTitleRow, checkTitleCellColumn];
                                }
                                else
                                {
                                    tempCheckNorm = thisSheetTemp.Cells[lastRowTemp, checkTitleCellColumn];
                                }
                                if (basicRules.Contains("Проверка на выравнивание текста"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].HorizontalAlignment != tempCheckNorm.HorizontalAlignment)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].HorizontalAlignment = tempCheckNorm.HorizontalAlignment;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm +" Ошибка выравнивания;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].VerticalAlignment != thisSheetTemp.Cells[lastRowTemp, checkTitleCellColumn].VerticalAlignment)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].VerticalAlignment = thisSheetTemp.Cells[lastRowTemp, checkTitleCellColumn].VerticalAlignment;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Ошибка выравнивания";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на размер шрифта"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Size != tempCheckNorm.Font.Size)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Size = tempCheckNorm.Font.Size;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Неверный размер шрифта;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на стиль шрифта"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Name != tempCheckNorm.Font.Name)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Name = tempCheckNorm.Font.Name;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Неверный стиль шрифта;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на жирность текста"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Bold != tempCheckNorm.Font.Bold)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Bold = tempCheckNorm.Font.Bold;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Неверная жирность текста;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на курсив текста"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Italic != tempCheckNorm.Font.Italic)
                                    {
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Font.Italic = tempCheckNorm.Font.Italic;
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Неверный курсив текста;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на двойные пробелы"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Text.Contains("  "))
                                    {
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Двойной пробел;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                                if (basicRules.Contains("Проверка на двойные знаки препинания"))
                                {
                                    if (thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Text.Contains("..")|
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(",,")|
                                        thisSheetDoc.Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(";;"))
                                    {
                                        string oldComm = "";
                                        if (thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + " Двойной знак препинания;";
                                        thisSheetDoc.Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    }
                                }
                            }

                        }
                    }
                }
            }
        }
        private void checkUserRules(Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {
            checkingUserRules = true;
            int i = 0;
            currentRow = 0;
            while (ruleFile[i].Contains("**Rules**") != true & i < ruleFile.Length - 2)
            {
                if (ruleFile[i + 2].Contains("checkState_YES") == true)
                {
                    if (ruleFile[i + 3].Contains("**??Rule??**") == false)
                    {
                        backgroundWorker1.ReportProgress(currentRow);
                        if (ruleFile[i + 3] == "shallInclude")
                        {
                            string condition = "";
                            int t = i + 4;
                            while (ruleFile[t] != "**??Rule??**")
                            {
                                if (ruleFile[t] == "")
                                {
                                    condition = condition + "**??";
                                }
                                else
                                {
                                    if (condition == "")
                                    {
                                        condition = ruleFile[t];
                                    }
                                    else
                                    {
                                        condition = condition + "**??" + ruleFile[t];
                                    }
                                }
                                t++;
                            }
                            shallInclude(ruleFile[i + 1], condition, docWorkBook);
                        }
                        if (ruleFile[i + 3] == "shallNotInclude")
                        {
                            string condition = "";
                            int t = i + 4;
                            while (ruleFile[t] != "**??Rule??**")
                            {
                                if (ruleFile[t] == "")
                                {
                                    condition = condition + "**??";
                                }
                                else
                                {
                                    if (condition == "")
                                    {
                                        condition = ruleFile[t];
                                    }
                                    else
                                    {
                                        condition = condition + "**??" + ruleFile[t];
                                    }
                                }
                                t++;
                            }
                            shallNotInclude(ruleFile[i + 1], condition, docWorkBook);
                        }
                        if (ruleFile[i + 3] == "if_A_then_B")
                        {
                            string condition = "";
                            int t = i + 5;
                            while (ruleFile[t] != "**??Rule??**")
                            {
                                if (ruleFile[t] == "")
                                {
                                    condition = condition + "**??";
                                }
                                else
                                {
                                    if (condition == "")
                                    {
                                        condition = ruleFile[t];
                                    }
                                    else
                                    {
                                        condition = condition + "**??" + ruleFile[t];
                                    }
                                }
                                t++;
                            }
                            if_A_then_B(ruleFile[i + 1], ruleFile[i + 4], condition, docWorkBook);
                        }
                        if (ruleFile[i + 3] == "if_A_not_B")
                        {
                            string condition = "";
                            int t = i + 5;
                            while (ruleFile[t] != "**??Rule??**")
                            {
                                if (ruleFile[t] == "")
                                {
                                    condition = condition + "**??";
                                }
                                else
                                {
                                    if (condition == "")
                                    {
                                        condition = ruleFile[t];
                                    }
                                    else
                                    {
                                        condition = condition + "**??" + ruleFile[t];
                                    }
                                }
                                t++;
                            }
                            if_A_not_B(ruleFile[i + 1], ruleFile[i + 4], condition, docWorkBook);
                        }
                        if (ruleFile[i + 3] == "if_A_then_A_in_B")
                        {
                            string condition = "";
                            int t = i + 5;
                            while (ruleFile[t] != "**??Rule??**")
                            {
                                if (ruleFile[t] == "")
                                {
                                    condition = condition + "**??";
                                }
                                else
                                {
                                    if (condition == "")
                                    {
                                        condition = ruleFile[t];
                                    }
                                    else
                                    {
                                        condition = condition + "**??" + ruleFile[t];
                                    }
                                }
                                t++;
                            }
                            if_A_then_A_in_B(ruleFile[i + 1], ruleFile[i + 4], condition, docWorkBook);
                        }
                        currentRow++;
                        backgroundWorker1.ReportProgress(currentRow);
                    }
                }
                i++;
            }
        }
        private void shallInclude(string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {
            using (StreamWriter sw = new StreamWriter(starting_page_Excel.fileForCheck.Substring(0, starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "\\" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf("\\"), starting_page_Excel.fileForCheck.LastIndexOf(".") - starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                int sheetsInExcel = docWorkBook.Sheets.Count;
                int y = 0;
                bool uu = false;
                while (y < subsub.Length)
                {
                    int ss = 0;
                    while (ss < sheetsInExcel & uu == false)
                    {
                        ss++;
                        var lastCell = docWorkBook.Sheets[ss].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                        int checkTitleRow = 0;
                        while (checkTitleRow < lastCell.Row & uu == false)
                        {
                            checkTitleRow++;
                            int checkTitleCellColumn = 0;
                            while (checkTitleCellColumn < lastCell.Column & uu == false)
                            {
                                checkTitleCellColumn++;
                                if (docWorkBook.Sheets[ss].Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(subjectOfSearch))
                                {
                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                    uu = true;
                                }
                            }
                        }
                    }
                    y++;
                }
                if (uu == false)
                {
                    sw.WriteLine("не нашлось (должно быть): " + "\"" + subsub[y] + "\"");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void shallNotInclude(string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {
            using (StreamWriter sw = new StreamWriter(starting_page_Excel.fileForCheck.Substring(0, starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "\\" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf("\\"), starting_page_Excel.fileForCheck.LastIndexOf(".") - starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                int sheetsInExcel = docWorkBook.Sheets.Count;
                int y = 0;
                bool uu = false;
                while (y < subsub.Length)
                {
                    int ss = 0;
                    while (ss < sheetsInExcel & uu == false)
                    {
                        ss++;
                        var lastCell = docWorkBook.Sheets[ss].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                        int checkTitleRow = 0;
                        while (checkTitleRow < lastCell.Row & uu == false)
                        {
                            checkTitleRow++;
                            int checkTitleCellColumn = 0;
                            while (checkTitleCellColumn < lastCell.Column & uu == false)
                            {
                                checkTitleCellColumn++;
                                if (docWorkBook.Sheets[ss].Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(subjectOfSearch))
                                {
                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                    string oldComm = "";
                                    if (docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                    {
                                        oldComm = docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                    }
                                    docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                    var Text = oldComm + "Не должно быть " + "\"" + subjectOfSearch + "\";";
                                    docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                    uu = true;
                                }
                            }
                        }
                    }
                    y++;
                }
                if (uu == false)
                {
                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_B(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {
            
        }
        private void if_A_not_B(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {

        }
        private void if_A_then_A_in_B(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Excel.Workbook docWorkBook)
        {
            using (StreamWriter sw = new StreamWriter(starting_page_Excel.fileForCheck.Substring(0, starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "\\" + starting_page_Excel.fileForCheck.Substring(starting_page_Excel.fileForCheck.LastIndexOf("\\"), starting_page_Excel.fileForCheck.LastIndexOf(".") - starting_page_Excel.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                int sheetsInExcel = docWorkBook.Sheets.Count;
                int y = 0;
                bool uu = true;
                int ss = 0;
                while (ss < sheetsInExcel)
                {
                    ss++;
                    var lastCell = docWorkBook.Sheets[ss].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    int checkTitleRow = 0;
                    while (checkTitleRow < lastCell.Row )
                    {
                        checkTitleRow++;
                        int checkTitleCellColumn = 0;
                        while (checkTitleCellColumn < lastCell.Column )
                        {
                            checkTitleCellColumn++;
                            if (docWorkBook.Sheets[ss].Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(subjectOfSearch_A))
                            {
                                int ret = 0;
                                bool pot = false;
                                while (ret < subsub.Length)
                                {
                                    if (docWorkBook.Sheets[ss].Cells[checkTitleRow, checkTitleCellColumn].Text.Contains(subsub[ret]) == true)
                                    {
                                        pot = true;
                                    }
                                    ret++;
                                    if (pot == false)
                                    {
                                        sw.WriteLine("найдено (не должно быть): " + "\"" + subjectOfSearch_A + "\"");
                                        string oldComm = "";
                                        if (docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment != null)
                                        {
                                            oldComm = docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].Comment.Shape.TextFrame.Characters().Text;
                                        }
                                        docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].ClearComments();
                                        var Text = oldComm + "Не должно быть " + "\"" + subjectOfSearch_A + "\";";
                                        docWorkBook.Sheets[ss].Range[NumberToLetters(checkTitleCellColumn) + checkTitleRow.ToString()].AddComment(Text);
                                        uu = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (uu == false)
                {
                    sw.WriteLine("                      ОК: ");
                }
                sw.WriteLine("---------------------------");
            }
        }
        static string NumberToLetters(int number)
        {
            string result;
            if (number > 0)
            {
                int alphabets = (number - 1) / 26;
                int remainder = (number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                    result = NumberToLetters(alphabets) + result;
            }
            else
                result = null;
            return result;
        }
        static int LettersToNumber(string letters)
        {
            int result = 0;
            if (letters.Length > 0 && letters.All(a => (a >= 'A' && a <= 'Z')))
                try
                {
                    for (int i = letters.Length; i > 0; i--)
                        result += (int)checked(Math.Pow(26, i - 1) + (letters[i - 1] - 'A') *
                            Math.Pow(26, letters.Length - i));
                }
                catch (OverflowException)
                {
                    result = -1;
                }
            else
                result = -1;
            return result;
        }
    }
}
