using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace Нормоконтроль
{
    public partial class Прогресс_проверки : Form
    {
        public Прогресс_проверки()
        {
            InitializeComponent();
            ruleFile = System.IO.File.ReadAllLines(Form1.rulesFile, Encoding.UTF8);
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
        Microsoft.Office.Interop.Word.Application appTemp;
        Microsoft.Office.Interop.Word.Document template;
        Microsoft.Office.Interop.Word.Application appDoc;
        Microsoft.Office.Interop.Word.Document doc;
        private void Прогресс_проверки_Load(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync();
            }

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (Form1.templateForCheck.Length>1)
            {
                appTemp = new Microsoft.Office.Interop.Word.Application();
                template = appTemp.Documents.Open(Form1.templateForCheck);
            }
            File.Copy(Form1.fileForCheck, Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_temp" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf(".")));
            
            appDoc = new Microsoft.Office.Interop.Word.Application();
            doc = appDoc.Documents.Open(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_temp" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf(".")));

            
            numberOfTables = doc.Tables.Count;
            numberOfParagraphInDoc = doc.Paragraphs.Count;
            int i = 0;
            while (i < doc.Paragraphs.Count) //подсчет абзацев вне таблиц (колонтитул?)
            {
                i++;
                if (doc.Paragraphs[i].Range.Tables.Count == 0 & doc.Paragraphs[i].Range.Text.Length > 1)
                {
                    numberOfParagraphOutOfTables++;
                }
            }
            i = 0;
            while (i < ruleFile.Length-1) //подсчет выбранных правил юзерка
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
            if (Form1.templateForCheck.Length > 1)
            {
                checkFormatingTables(doc, template);
            }
            paragraphOutOfTables(doc, template);
            checkUserRules(doc);

            
            doc.SaveAs2(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_после проверки" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf(".")));
            doc.Close();
            appDoc.Quit();
            File.Delete(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_temp" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf(".")));
            if (Form1.templateForCheck.Length > 1)
            {
                appTemp.ActiveDocument.Close();
                appTemp.Quit();
            }

        }
        private void checkFormatingTables(Microsoft.Office.Interop.Word.Document doc, Microsoft.Office.Interop.Word.Document template)
        {
            foreach (Word.Table Table in doc.Tables)
            {
                int i = 0;
                bool b = false;
                while (b != true & i < template.Tables.Count)
                {
                    i++;
                    if (template.Tables[i].Rows[1].Cells.Count == Table.Rows[1].Cells.Count)
                    {
                        bool q = true;
                        int thisColumnNumber = 0;
                        while (q == true & thisColumnNumber < Table.Rows[1].Cells.Count)
                        {
                            thisColumnNumber++;
                            if (Table.Cell(1, thisColumnNumber).Range.Text == template.Tables[i].Cell(1, thisColumnNumber).Range.Text)
                            {
                                q = true;
                            }
                            else
                            {
                                q = false;
                            }
                            if (thisColumnNumber == Table.Rows[1].Cells.Count & q == true)
                            {
                                b = true;
                            }
                        }
                    }
                }
                if (b == true)
                {
                    checkingTables = true;
                    int numberTrueTableInTemplate = i;
                    currentTableFormZero++;
                    numberOfRowsInCurrentTable = Table.Rows.Count;
                    i = 0;
                    while (i < Table.Rows.Count)
                    {
                        
                        i++;
                        currentRow = i;
                        backgroundWorker1.ReportProgress(currentRow);
                        int iInTemplate = i;
                        int numberOfCellInRow = 0;
                        if (i > template.Tables[numberTrueTableInTemplate].Rows.Count)
                        {
                            iInTemplate = template.Tables[numberTrueTableInTemplate].Rows.Count;
                        }
                        while (numberOfCellInRow < Table.Rows[i].Cells.Count)
                        {
                            numberOfCellInRow++;
                            if (basicRules.Contains("Проверка на выравнивание текста"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.ParagraphFormat.Alignment != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.ParagraphFormat.Alignment)
                                {
                                    Table.Cell(i, numberOfCellInRow).Range.ParagraphFormat.Alignment = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.ParagraphFormat.Alignment;
                                    object text = "Неверное выравнивание";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на размер шрифта"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Font.Size != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Size)
                                {
                                    Table.Cell(i, numberOfCellInRow).Range.Font.Size = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Size;
                                    object text = "Неверный размер шрифта";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на стиль шрифта"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Font.Name != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Name)
                                {
                                    Table.Cell(i, numberOfCellInRow).Range.Font.Name = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Name;
                                    object text = "Неверный стиль шрифта";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на жирность текста"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Font.Bold != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Bold)
                                {
                                    Table.Cell(i, numberOfCellInRow).Range.Font.Bold = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Bold;
                                    object text = "Жирность";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на курсив текста"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Font.Italic != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Italic)
                                {
                                    Table.Cell(i, numberOfCellInRow).Range.Font.Italic = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Font.Italic;
                                    object text = "Курсив";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на двойные пробелы"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Text.Contains("  "))
                                {
                                    object text = "Двойной пробел";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }

                            if (basicRules.Contains("Проверка на двойные знаки препинания"))
                            {
                                if (Table.Cell(i, numberOfCellInRow).Range.Text.Contains("..") | 
                                    Table.Cell(i, numberOfCellInRow).Range.Text.Contains(",,") |
                                    Table.Cell(i, numberOfCellInRow).Range.Text.Contains(";;") |
                                    Table.Cell(i, numberOfCellInRow).Range.Text.Contains("::") |
                                    Table.Cell(i, numberOfCellInRow).Range.Text.Contains("!!") |
                                    Table.Cell(i, numberOfCellInRow).Range.Text.Contains("??"))
                                {
                                    object text = "Двойной знак препинания";
                                    try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                }
                            }
                            if (basicRules.Contains("Проверка: цвет текста везде черный"))
                            {
                                int oo = 0;
                                while (oo < Table.Cell(i, numberOfCellInRow).Range.Words.Count)
                                {
                                    oo++;
                                    if (Table.Cell(i, numberOfCellInRow).Range.Words[oo].Font.Color != template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Words[1].Font.Color)
                                    {
                                        Table.Cell(i, numberOfCellInRow).Range.Words[oo].Font.Color = template.Tables[numberTrueTableInTemplate].Cell(iInTemplate, numberOfCellInRow).Range.Words[1].Font.Color;
                                        object text = "Цвет текста не автоматический (по умолчанию - черный)";
                                        try { doc.Comments.Add(Table.Cell(i, numberOfCellInRow).Range, ref text); } catch { }
                                        break;
                                    }
                                }

                            }
                        }
                    }
                    checkingTables = false;
                }
            }
            checkedTables.ForeColor = Color.Black;
        }
        private void paragraphOutOfTables (Microsoft.Office.Interop.Word.Document doc, Microsoft.Office.Interop.Word.Document template)
        {
            checkingParagraphsOutOfTables = true;
            string textStyle = "", textSize = "", bold = "", indentRight = "", indentFirstRow = "", intervalBefore = "", intervalAfter = "";
            int inTempI = 0;
            int i = 0;
            while (i< ruleFile.Length) //поиск инфы о абзацах
            {
                if (ruleFile[i].Contains("WO paragraph") == true)
                {
                    paraInTempChosed = true;
                    int ii = 0;
                    if (template != null)
                    {
                        while (ii != template.Paragraphs.Count)
                        {
                            ii++;
                            if (template.Paragraphs[ii].Range.Tables.Count == 0 & template.Paragraphs[ii].Range.Text.Length > 200)
                            {
                                textStyle = template.Paragraphs[ii].Range.Words[1].Font.Name.ToString();
                                textSize = template.Paragraphs[ii].Range.Words[1].Font.Size.ToString();
                                textAlighnment = template.Paragraphs[ii].Range.ParagraphFormat.Alignment.ToString();
                                italic = template.Paragraphs[ii].Range.Words[1].Font.Italic.ToString();
                                bold = template.Paragraphs[ii].Range.Words[1].Font.Bold.ToString();
                                indentLeft = template.Paragraphs[ii].Range.ParagraphFormat.LeftIndent.ToString();
                                indentRight = template.Paragraphs[ii].Range.ParagraphFormat.RightIndent.ToString();
                                indentFirstRow = template.Paragraphs[ii].Range.ParagraphFormat.FirstLineIndent.ToString();
                                intervalBefore = template.Paragraphs[ii].Range.ParagraphFormat.SpaceBefore.ToString();
                                intervalAfter = template.Paragraphs[ii].Range.ParagraphFormat.SpaceAfter.ToString();
                                paraInTempExsistents = true;
                                inTempI = ii;
                            }
                        }
                    }
                }
                if (ruleFile[i] == "*Formating without paragraph template*")
                {
                    if (ruleFile[i+1].Contains("WO paragraph template") == false)
                    {
                        textStyle = ruleFile[i + 1];
                        textSize = ruleFile[i + 2];
                        //
                        if (ruleFile[i + 3] == "Слева")
                        {
                            textAlighnment = "wdAlignParagraphLeft";
                        }
                        if (ruleFile[i + 3] == "Справа")
                        {
                            textAlighnment = "wdAlignParagraphRight";
                        }
                        if (ruleFile[i + 3] == "По центру")
                        {
                            textAlighnment = "wdAlignParagraphCenter";
                        }
                        if (ruleFile[i + 3] == "По ширине")
                        {
                            textAlighnment = "wdAlignParagraphJustify";
                        }
                        //
                        if (ruleFile[i + 4] == "Да")
                        {
                            italic = "1";
                        }
                        if (ruleFile[i + 4] == "Нет")
                        {
                            italic = "0";
                        }
                        //
                        if (ruleFile[i + 5] == "Да")
                        {
                            bold = "1";
                        }
                        if (ruleFile[i + 5] == "Нет")
                        {
                            bold = "0";
                        }
                        //
                        indentLeft = ruleFile[i + 6];
                        indentRight = ruleFile[i + 7];
                        indentFirstRow = ruleFile[i + 8];
                        intervalBefore = ruleFile[i + 9];
                        intervalAfter = ruleFile[i + 10];
                        break;
                    }
                }
                i++;
            }
            if (paraInTempChosed == true & paraInTempExsistents != true)
            {
                checkingParagraphsOutOfTables = false;
                checkedFormating.ForeColor = Color.Silver;
                return;
            }
            currentRow = 0;
            if (paraInTempExsistents == false)
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    if (para.Range.Tables.Count == 0 & para.Range.Text.Length > 1)
                    {
                        currentRow++;
                        backgroundWorker1.ReportProgress(currentRow);
                        if (basicRules.Contains("Проверка на стиль шрифта"))
                        {
                            if (para.Range.Font.Name.ToString() != textStyle & textStyle.Length > 0)
                            {
                                object text = "Неверный стиль шрифта";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на размер шрифта"))
                        {
                            if (para.Range.Font.Size.ToString() != textSize & textSize.Length > 0)
                            {
                                object text = "Неверный размер шрифта";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на курсив текста"))
                        {
                            if (para.Range.Font.Italic.ToString() != italic & italic.Length > 0)
                            {
                                object text = "Курсив";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на жирность текста"))
                        {
                            if (para.Range.Font.Bold.ToString() != bold & bold.Length > 0)
                            {
                                object text = "Жирность текста";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на выравнивание текста"))
                        {
                            if (para.Range.ParagraphFormat.Alignment.ToString() != textAlighnment & textAlighnment.Length > 0)
                            {
                                object text = "Неверное выравнивание текста";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.LeftIndent.ToString() != indentLeft & indentLeft.Length > 0)
                            {
                                object text = "Неверный отступ слева";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.RightIndent.ToString() != indentRight & indentRight.Length > 0)
                            {
                                object text = "Неверный отступ справа";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.FirstLineIndent.ToString() != indentFirstRow & indentFirstRow.Length > 0)
                            {
                                object text = "Неверный отступ первой строки";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на интервалы до/после"))
                        {
                            if (para.Range.ParagraphFormat.SpaceBefore.ToString() != intervalBefore & intervalBefore.Length > 0)
                            {
                                object text = "Неверный интервал до абзаца";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на интервалы до/после"))
                        {
                            if (para.Range.ParagraphFormat.SpaceAfter.ToString() != intervalAfter & intervalAfter.Length > 0)
                            {
                                object text = "Неверный интервал после абзаца";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на двойные пробелы"))
                        {
                            if (para.Range.Text.Contains("  "))
                            {
                                object text = "Двойной пробел";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на двойные знаки препинания"))
                        {
                            if (para.Range.Text.Contains("..") |
                                para.Range.Text.Contains(",,") |
                                para.Range.Text.Contains("::") |
                                para.Range.Text.Contains("!!") |
                                para.Range.Text.Contains("??"))
                            {
                                object text = "Двойной знак препинания";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка: цвет текста везде черный"))
                        {
                            int oo = 0;
                            while (oo < para.Range.Words.Count)
                            {
                                oo++;
                                if (para.Range.Words[oo].Font.Color.ToString() != "wdColorAutomatic")
                                {
                                    object text = "Цвет текста не автоматический (по умолчанию - черный)";
                                    try { doc.Comments.Add(para.Range, ref text); } catch { }
                                    break;
                                }
                            }

                        }
                        /*if (basicRules.Contains("Проверка: цвет фона текста везде отсутствует"))
                        {
                            int oo = 0;
                            while (oo < para.Range.Words.Count)
                            {
                                oo++;
                                if (para.Range.Words[oo].Font.Shading.ForegroundPatternColor.ToString() != "wdColorAutomatic")
                                {
                                    object text = "Цвет фона текста не белый";
                                    try { doc.Comments.Add(para.Range, ref text); } catch { }
                                    break;
                                }
                            }

                        }*/
                    }
                }
            }
            if (paraInTempExsistents != false)
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    if (para.Range.Tables.Count == 0 & para.Range.Text.Length > 1)
                    {
                        currentRow++;
                        backgroundWorker1.ReportProgress(currentRow);
                        if (basicRules.Contains("Проверка на стиль шрифта"))
                        {
                            if (para.Range.Font.Name.ToString() != textStyle & textStyle.Length > 0)
                            {
                                object text = "Неверный стиль шрифта";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                                para.Range.Font.Name = template.Paragraphs[inTempI].Range.Words[1].Font.Name;
                            }
                        }
                        if (basicRules.Contains("Проверка на размер шрифта"))
                        {
                            if (para.Range.Font.Size.ToString() != textSize & textSize.Length > 0)
                            {
                                para.Range.Font.Size = template.Paragraphs[inTempI].Range.Words[1].Font.Size;
                                object text = "Неверный размер шрифта";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на курсив текста"))
                        {
                            if (para.Range.Font.Italic.ToString() != italic & italic.Length > 0)
                            {
                                para.Range.Font.Italic = template.Paragraphs[inTempI].Range.Words[1].Font.Italic;
                                object text = "Курсив";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на жирность текста"))
                        {
                            if (para.Range.Font.Bold.ToString() != bold & bold.Length > 0)
                            {
                                para.Range.Font.Bold = template.Paragraphs[inTempI].Range.Words[1].Font.Bold;
                                object text = "Жирность текста";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на выравнивание текста"))
                        {
                            if (para.Range.ParagraphFormat.Alignment.ToString() != textAlighnment & textAlighnment.Length > 0)
                            {
                                para.Range.ParagraphFormat.Alignment = template.Paragraphs[inTempI].Range.ParagraphFormat.Alignment;
                                object text = "Неверное выравнивание текста";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.LeftIndent.ToString() != indentLeft & indentLeft.Length > 0)
                            {
                                para.Range.ParagraphFormat.LeftIndent = template.Paragraphs[inTempI].Range.ParagraphFormat.LeftIndent;
                                object text = "Неверный отступ слева";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.RightIndent.ToString() != indentRight & indentRight.Length > 0)
                            {
                                para.Range.ParagraphFormat.RightIndent = template.Paragraphs[inTempI].Range.ParagraphFormat.RightIndent;
                                object text = "Неверный отступ справа";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на отступы слева/справа/первой строки"))
                        {
                            if (para.Range.ParagraphFormat.FirstLineIndent.ToString() != indentFirstRow & indentFirstRow.Length > 0)
                            {
                                para.Range.ParagraphFormat.FirstLineIndent = template.Paragraphs[inTempI].Range.ParagraphFormat.FirstLineIndent;
                                object text = "Неверный отступ первой строки";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на интервалы до/после"))
                        {
                            if (para.Range.ParagraphFormat.SpaceBefore.ToString() != intervalBefore & intervalBefore.Length > 0)
                            {
                                para.Range.ParagraphFormat.SpaceBefore = template.Paragraphs[inTempI].Range.ParagraphFormat.SpaceBefore;
                                object text = "Неверный интервал до абзаца";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на интервалы до/после"))
                        {
                            if (para.Range.ParagraphFormat.SpaceAfter.ToString() != intervalAfter & intervalAfter.Length > 0)
                            {
                                para.Range.ParagraphFormat.SpaceAfter = template.Paragraphs[inTempI].Range.ParagraphFormat.SpaceAfter;
                                object text = "Неверный интервал после абзаца";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на двойные пробелы"))
                        {
                            if (para.Range.Text.Contains("  "))
                            {
                                object text = "Двойной пробел";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка на двойные знаки препинания"))
                        {
                            if (para.Range.Text.Contains("..") |
                                para.Range.Text.Contains(",,") |
                                para.Range.Text.Contains("::") |
                                para.Range.Text.Contains("!!") |
                                para.Range.Text.Contains("??"))
                            {
                                object text = "Двойной знак препинания";
                                try { doc.Comments.Add(para.Range, ref text); } catch { }
                            }
                        }
                        if (basicRules.Contains("Проверка: цвет текста везде черный"))
                        {
                            int oo = 0;
                            while (oo < para.Range.Words.Count)
                            {
                                oo++;
                                if (para.Range.Words[oo].Font.Color != template.Paragraphs[inTempI].Range.Words[1].Font.Color)
                                {
                                    para.Range.Words[oo].Font.Color = template.Paragraphs[inTempI].Range.Words[1].Font.Color;
                                    object text = "Цвет текста не по шаблону";
                                    try { doc.Comments.Add(para.Range, ref text); } catch { }
                                    break;
                                }
                            }

                        }
                        /*if (basicRules.Contains("Проверка: цвет фона текста везде отсутствует"))
                        {
                            int oo = 0;
                            while (oo < para.Range.Words.Count)
                            {
                                oo++;
                                if (para.Range.Words[oo].Font.Shading.ForegroundPatternColor.ToString() != "wdColorAutomatic")
                                {
                                    object text = "Цвет фона текста не белый";
                                    try { doc.Comments.Add(para.Range, ref text); } catch { }
                                    break;
                                }
                            }

                        }*/
                    }
                }
            }
            checkedFormating.ForeColor = Color.Black;
            checkingParagraphsOutOfTables = false;
        }
        private void checkUserRules(Microsoft.Office.Interop.Word.Document doc)
        {
            checkingUserRules = true;
            int i = 0;
            currentRow = 0;
            while (ruleFile[i].Contains("**Rules**") != true & i<ruleFile.Length-2)
            {
                if (ruleFile[i+2].Contains("checkState_YES") == true)
                {
                    if (ruleFile[i + 3].Contains("**??Rule??**") == false)
                    {
                        backgroundWorker1.ReportProgress(currentRow);
                        if (ruleFile[i + 3] == "shallInclude")
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                shallInclude(ruleFile[i + 1], condition, doc);
                            else 
                                shallInclude_withRunningTitles(ruleFile[i + 1], condition, doc);
                        }
                        if (ruleFile[i + 3] == "shallNotInclude")
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                shallNotInclude(ruleFile[i + 1], condition, doc);
                            else
                                shallNotInclude_withRunningTitles(ruleFile[i + 1], condition, doc);
                        }
                        if (ruleFile[i + 3] == "if_A_then_B")
                        {
                            string condition = "";
                            int t = i + 6;
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                if_A_then_B(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                            else
                                if_A_then_B_withRunningTitles(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                        }
                        if (ruleFile[i + 3] == "if_A_not_B")
                        {
                            string condition = "";
                            int t = i + 6;
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                if_A_not_B(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                            else
                                if_A_not_B_withRunningTitles(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                        }
                        if (ruleFile[i + 3] == "if_A_then_A_in_B")
                        {
                            string condition = "";
                            int t = i + 6;
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                if_A_then_A_in_B(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                            else
                                if_A_then_A_in_B_withRunningTitles(ruleFile[i + 1], ruleFile[i + 5], condition, doc);
                        }
                        if (ruleFile[i + 3] == "if_A_then_B_multistringA")
                        {
                            string condition = "";
                            int t = i + 5;
                            while (ruleFile[t+1] != "**??Rule??**")
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
                            if (ruleFile[i + 4] != "runningTitle_YES")
                                if_A_then_B_multistringA(ruleFile[i + 1], condition, ruleFile[t], doc);
                            else
                                if_A_then_B_multistringA_withRunningTitles(ruleFile[i + 1], condition, ruleFile[t], doc);
                        }
                        currentRow++;
                        backgroundWorker1.ReportProgress(currentRow);
                    }
                }
                i++;
            }
        }
        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
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
        }
        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            checkedRules.ForeColor = Color.Black;
            checkingIsDone.Visible = true;
            commentsForFinal.Visible = true;
        }
        private void Прогресс_проверки_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Visible = true;
            try
            {
                appDoc.Quit();
            }
            catch { }
            if (appTemp!=null)
            try
            {
                appTemp.Quit();
            }
            catch { }
        }
        private void shallInclude(string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = true;
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                            uu = true;
                        }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        
                        if (uu != true)
                        {
                            sw.WriteLine("не нашлось (должно быть): " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void shallNotInclude (string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = true;
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                            uu = true;
                            int ip = 0;
                            while (ip < doc.Paragraphs.Count)
                            {
                                ip++;
                                if (doc.Paragraphs[ip].Range.Text.Contains(subsub[y]))
                                {
                                    object text = "не должно быть: " + "\"" + subsub[y] + "\"";
                                    try { doc.Comments.Add(doc.Paragraphs[ip].Range, ref text); } catch { }
                                }
                            }
                        }
                        
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        
                        if (uu != true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_B (string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                }
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                            uu = true;
                        }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        
                        if (uu != true)
                        {
                            sw.WriteLine("не нашлось (должно быть): " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                else
                {
                    sw.WriteLine("исходное условие \"" + subjectOfSearch_A + "\" не нашлось");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_B_multistringA(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                bool uu = false;
                string[] subsub = subjectOfSearch_A.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                int y = 0;
                while (y < subsub.Length)
                {
                    if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                    {
                        uu = true;
                        if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                        {
                            findIn = true;
                            sw.WriteLine("                      ОК: ");
                            break;
                        }
                    }
                    y++;
                }

                /*if (uu != true)
                {
                    int yy = 1;
                    y = 0;
                    while (yy <= doc.Shapes.Count | findIn != true)
                    {
                        while (y < subsub.Length | findIn != true)
                        {
                            try
                            {
                                if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                {
                                    uu = true;
                                    try
                                    {
                                        if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_B) == true)
                                        {
                                            findIn = true;
                                            sw.WriteLine("                      ОК: ");
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                            y++;
                        }
                        yy++;
                    }
                }*/

                if (uu != true)
                {
                    sw.WriteLine("исходное условие не нашлось");
                }
                if (findIn != true & uu == true)
                {
                    sw.WriteLine("не нашлось (должно быть): " + "\"" + subjectOfSearch_B + "\"");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_not_B(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                }
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                            uu = true;
                            int ip = 0;
                            while (ip < doc.Paragraphs.Count)
                            {
                                ip++;
                                if (doc.Paragraphs[ip].Range.Text.Contains(subsub[y]))
                                {
                                    object text = "не должно быть: " + "\"" + subsub[y] + "\"";
                                    try { doc.Comments.Add(doc.Paragraphs[ip].Range, ref text); } catch { }
                                }
                            }
                        }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (uu != true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                else
                {
                    sw.WriteLine("исходное условие \"" + subjectOfSearch_A + "\" не нашлось");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_A_in_B(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                string wholeWord = "";
                string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                    int i = 0;
                    while (i<doc.Paragraphs.Count)
                    {
                        i++;
                        if (doc.Paragraphs[i].Range.Text.Contains(subjectOfSearch_A))
                        {
                            int ret = 0;
                            bool pot = false;
                            while (ret<subsub.Length)
                            {
                                if (doc.Paragraphs[i].Range.Text.Contains(subsub[ret])==true)
                                {
                                    pot = true;
                                }
                                    
                                ret++;
                            }
                            if (pot==false)
                            {
                                sw.WriteLine("найдено (не должно быть): " + "\"" + subjectOfSearch_A + "\"");

                                object text = "не должно быть: " + "\"" + subjectOfSearch_A + "\"";

                                try { doc.Comments.Add(doc.Paragraphs[i].Range, ref text); } catch { }

                            }
                        }
                    }
                }
                if (findIn != true)
                {
                    sw.WriteLine("                      ОК: ");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void shallInclude_withRunningTitles(string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = true;
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                            uu = true;
                        }
                        try
                        {
                            if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == false)
                            {
                                foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                                {
                                    {
                                        if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y])) & uu != true)
                                        {
                                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                            uu = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (uu != true)
                        {
                            foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                            {
                                if (uu != true)
                                {
                                    int yy = 0;
                                    while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                    {
                                        yy++;
                                        try
                                        {
                                            if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                            {
                                                sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                uu = true;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                        }
                        if (uu != true)
                        {
                            sw.WriteLine("не нашлось (должно быть): " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void shallNotInclude_withRunningTitles(string nameOfRule, string subjectOfSearch, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = true;
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                            uu = true;
                            int ip = 0;
                            while (ip < doc.Paragraphs.Count)
                            {
                                ip++;
                                if (doc.Paragraphs[ip].Range.Text.Contains(subsub[y]))
                                {
                                    object text = "не должно быть: " + "\"" + subsub[y] + "\"";
                                    try { doc.Comments.Add(doc.Paragraphs[ip].Range, ref text); } catch { }
                                }
                            }
                        }
                        try
                        {
                            if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == false)
                            {
                                foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                                {
                                    {
                                        if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y])) & uu != true)
                                        {
                                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                            uu = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (uu != true)
                        {
                            foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                            {
                                if (uu != true)
                                {
                                    int yy = 0;
                                    while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                    {
                                        yy++;
                                        try
                                        {
                                            if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                            {
                                                sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                uu = true;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                        }
                        if (uu != true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_B_withRunningTitles(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                }
                if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        {
                            if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A)) & findIn != true)
                            {
                                findIn = true;
                            }
                        }
                    }
                }
                if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        if (findIn != true)
                        {
                            int yy = 0;
                            while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                            {
                                yy++;
                                try
                                {
                                    if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                    {
                                        findIn = true;
                                    }
                                }
                                catch { }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                            uu = true;
                        }
                        try
                        {
                            if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == false)
                            {
                                foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                                {
                                    {
                                        if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y])) & uu != true)
                                        {
                                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                            uu = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (uu != true)
                        {
                            foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                            {
                                if (uu != true)
                                {
                                    int yy = 0;
                                    while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                    {
                                        yy++;
                                        try
                                        {
                                            if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                            {
                                                sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                uu = true;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                        }
                        if (uu != true)
                        {
                            sw.WriteLine("не нашлось (должно быть): " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                else
                {
                    sw.WriteLine("исходное условие \"" + subjectOfSearch_A + "\" не нашлось");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_not_B_withRunningTitles(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                }
                if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        {
                            if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A)) & findIn != true)
                            {
                                findIn = true;
                            }
                        }
                    }
                }
                if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        if (findIn != true)
                        {
                            int yy = 0;
                            while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                            {
                                yy++;
                                try
                                {
                                    if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                    {
                                        findIn = true;
                                    }
                                }
                                catch { }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
                if (findIn == true)
                {
                    string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                    int y = 0;
                    while (y < subsub.Length)
                    {
                        bool uu = false;
                        if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                        {
                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                            uu = true;
                            int ip = 0;
                            while (ip < doc.Paragraphs.Count)
                            {
                                ip++;
                                if (doc.Paragraphs[ip].Range.Text.Contains(subsub[y]))
                                {
                                    object text = "не должно быть: " + "\"" + subsub[y] + "\"";
                                    try { doc.Comments.Add(doc.Paragraphs[ip].Range, ref text); } catch { }
                                }
                            }
                        }
                        try
                        {
                            if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == false)
                            {
                                foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                                {
                                    {
                                        if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subsub[y]) |
                                            sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subsub[y])) & uu != true)
                                        {
                                            sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                            uu = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        if (uu != true)
                        {
                            int yy = 0;
                            while (yy < doc.Shapes.Count)
                            {
                                yy++;
                                try
                                {
                                    if (doc.Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                    {
                                        sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                        uu = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (uu != true)
                        {
                            foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                            {
                                if (uu != true)
                                {
                                    int yy = 0;
                                    while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                    {
                                        yy++;
                                        try
                                        {
                                            if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                            {
                                                sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                uu = true;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    if (uu != true)
                                    {
                                        yy = 0;
                                        while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & uu != true)
                                        {
                                            yy++;
                                            try
                                            {
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subsub[y]) == true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + subsub[y] + "\"");
                                                    uu = true;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                        }
                        if (uu != true)
                        {
                            sw.WriteLine("                      ОК: " + "\"" + subsub[y] + "\"");
                        }
                        y++;
                    }
                }
                else
                {
                    sw.WriteLine("исходное условие \"" + subjectOfSearch_A + "\" не нашлось");
                }
                sw.WriteLine("---------------------------");
            }
        }

        private void if_A_then_A_in_B_withRunningTitles(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                string wholeWord = "";
                string[] subsub = subjectOfSearch_B.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_A) == true)
                {
                    findIn = true;
                    int i = 0;
                    while (i < doc.Paragraphs.Count)
                    {
                        i++;
                        if (doc.Paragraphs[i].Range.Text.Contains(subjectOfSearch_A))
                        {
                            int ret = 0;
                            bool pot = false;
                            while (ret < subsub.Length)
                            {
                                if (doc.Paragraphs[i].Range.Text.Contains(subsub[ret]) == true)
                                {
                                    pot = true;
                                }

                                ret++;
                            }
                            if (pot == false)
                            {
                                sw.WriteLine("найдено (не должно быть): " + "\"" + subjectOfSearch_A + "\"");

                                object text = "не должно быть: " + "\"" + subjectOfSearch_A + "\"";

                                try { doc.Comments.Add(doc.Paragraphs[i].Range, ref text); } catch { }

                            }
                        }
                    }
                }
                /*if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        {
                            if ((sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Contains(subjectOfSearch_A) |
                                sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(subjectOfSearch_A)) & findIn != true)
                            {

                            }
                        }
                    }
                }*/
                if (findIn != true)
                {
                    foreach (Microsoft.Office.Interop.Word.Section sec in doc.Sections)
                    {
                        if (findIn != true)
                        {
                            int yy = 0;
                            while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                            {
                                yy++;
                                try
                                {
                                    if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                    {
                                        findIn = true;
                                        int i = 0;
                                        while (i < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words.Count)
                                        {
                                            i++;
                                            if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i].Text.Contains(subjectOfSearch_A))
                                            {
                                                wholeWord = sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i].Text;
                                                if (subjectOfSearch_B.Contains(wholeWord) != true)
                                                {
                                                    sw.WriteLine("найдено (не должно быть): " + "\"" + wholeWord + "\"");
                                                    object text = "не должно быть: " + "\"" + wholeWord + "\"";
                                                    try { doc.Comments.Add(sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i], ref text); } catch { }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch { }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                            int i = 0;
                                            while (i < sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words.Count)
                                            {
                                                i++;
                                                if (sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i].Text.Contains(subjectOfSearch_A))
                                                {
                                                    wholeWord = sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i].Text;
                                                    if (subjectOfSearch_B.Contains(wholeWord) != true)
                                                    {
                                                        sw.WriteLine("найдено (не должно быть): " + "\"" + wholeWord + "\"");
                                                        object text = "не должно быть: " + "\"" + wholeWord + "\"";
                                                        try { doc.Comments.Add(sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i], ref text); } catch { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                            int i = 0;
                                            while (i < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words.Count)
                                            {
                                                i++;
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i].Text.Contains(subjectOfSearch_A))
                                                {
                                                    wholeWord = sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i].Text;
                                                    if (subjectOfSearch_B.Contains(wholeWord) != true)
                                                    {
                                                        sw.WriteLine("найдено (не должно быть): " + "\"" + wholeWord + "\"");
                                                        object text = "не должно быть: " + "\"" + wholeWord + "\"";
                                                        try { doc.Comments.Add(sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[yy].TextFrame.TextRange.Words[i], ref text); } catch { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            if (findIn != true)
                            {
                                yy = 0;
                                while (yy < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count & findIn != true)
                                {
                                    yy++;
                                    try
                                    {
                                        if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Text.Contains(subjectOfSearch_A) == true)
                                        {
                                            findIn = true;
                                            int i = 0;
                                            while (i < sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words.Count)
                                            {
                                                i++;
                                                if (sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i].Text.Contains(subjectOfSearch_A))
                                                {
                                                    wholeWord = sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i].Text;
                                                    if (subjectOfSearch_B.Contains(wholeWord) != true)
                                                    {
                                                        sw.WriteLine("найдено (не должно быть): " + "\"" + wholeWord + "\"");
                                                        object text = "не должно быть: " + "\"" + wholeWord + "\"";
                                                        try { doc.Comments.Add(sec.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[yy].TextFrame.TextRange.Words[i], ref text); } catch { }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
                if (findIn != true)
                {
                    sw.WriteLine("                      ОК: ");
                }
                sw.WriteLine("---------------------------");
            }
        }
        private void if_A_then_B_multistringA_withRunningTitles(string nameOfRule, string subjectOfSearch_A, string subjectOfSearch_B, Microsoft.Office.Interop.Word.Document doc)
        {
            using (StreamWriter sw = new StreamWriter(Form1.fileForCheck.Substring(0, Form1.fileForCheck.LastIndexOf("\\")) + "\\" + Form1.fileForCheck.Substring(Form1.fileForCheck.LastIndexOf("\\"), Form1.fileForCheck.LastIndexOf(".") - Form1.fileForCheck.LastIndexOf("\\")) + "_log.txt", true))
            {
                sw.WriteLine(""); sw.WriteLine(""); sw.WriteLine("");
                sw.WriteLine(nameOfRule + ":");
                sw.WriteLine("---------------------------");
                object start = doc.Content.Start;
                object end = doc.Content.End;
                bool findIn = false;
                bool uu = false;
                string[] subsub = subjectOfSearch_A.Split(new[] { "**??" }, StringSplitOptions.RemoveEmptyEntries);
                int y = 0;
                while (y < subsub.Length)
                {
                    if (doc.Range(ref start, ref end).Text.Contains(subsub[y]) == true)
                    {
                        uu = true;
                        if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                        {
                            findIn = true;
                            sw.WriteLine("                      ОК: ");
                            break;
                        }
                    }
                    y++;
                }
                if (uu != true)
                {
                    y = 0;
                    while (y < subsub.Length)
                    {
                        foreach(Word.Shape sh in doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes)
                        {
                            try
                            {
                                if (sh.TextFrame.TextRange.Text.Contains(subsub[y]) & uu != true)
                                {
                                    uu = true;
                                    if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                                    {
                                        findIn = true;
                                        sw.WriteLine("                      ОК: ");
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        foreach (Word.Shape sh in doc.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                        {
                            try
                            {
                                if (sh.TextFrame.TextRange.Text.Contains(subsub[y]) & uu != true)
                                {
                                    uu = true;
                                    if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                                    {
                                        findIn = true;
                                        sw.WriteLine("                      ОК: ");
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        foreach (Word.Shape sh in doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes)
                        {
                            try
                            {
                                if (sh.TextFrame.TextRange.Text.Contains(subsub[y]) & uu != true)
                                {
                                    uu = true;
                                    if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                                    {
                                        findIn = true;
                                        sw.WriteLine("                      ОК: ");
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        foreach (Word.Shape sh in doc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                        {
                            try
                            {
                                if (sh.TextFrame.TextRange.Text.Contains(subsub[y]) & uu != true)
                                {
                                    uu = true;
                                    if (doc.Range(ref start, ref end).Text.Contains(subjectOfSearch_B) == true)
                                    {
                                        findIn = true;
                                        sw.WriteLine("                      ОК: ");
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        y++;
                    }
                        
                }
                if (uu != true)
                {
                    sw.WriteLine("исходное условие не нашлось");
                }
                if (findIn != true & uu == true)
                {
                    sw.WriteLine("не нашлось (должно быть): " + "\"" + subjectOfSearch_B + "\"");
                }
                sw.WriteLine("---------------------------");
            }
        }
    }
}
