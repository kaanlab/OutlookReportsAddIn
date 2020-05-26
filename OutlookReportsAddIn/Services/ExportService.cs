using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;


namespace OutlookReportsAddIn
{
    public class ExportService
    {
        public void ToWord(DateTime selectedDate, IEnumerable<Mail> emails, int counter = 1)
        {
            try
            {
                //Create an instance for word app
                Word.Application winword = new Word.Application
                {
                    //Set animation status for word application
                    ShowAnimation = false,
                    //Set status for word application is to be visible or not.
                    Visible = false

                };

                //Create a missing variable for missing value
                object missing = System.Reflection.Missing.Value;

                //Use ConfigurationHelper class to read OutlookReportsAddIn.dll.config 
                
                object filepath = Properties.Settings.Default.TemplatePath; 

                //Create a new document
                Word.Document document = winword.Documents.Add(ref filepath, ref missing, ref missing, ref missing);

                //Add paragraph 
                Word.Paragraph parag = document.Content.Paragraphs.Add(ref missing);
                parag.Range.InsertParagraphAfter();

                //Create new table in paragraph
                Word.Table table = document.Tables.Add(parag.Range, 3, 8, ref missing, ref missing);

                // Add border
                table.Borders.Enable = 1;

                // Add width for every colum
                table.Columns[1].Width = 28f;
                table.Columns[2].Width = 124f;
                table.Columns[3].Width = 180f;
                table.Columns[4].Width = 54f;
                table.Columns[5].Width = 44f;
                table.Columns[6].Width = 120f;
                table.Columns[7].Width = 100f;
                table.Columns[8].Width = 140f;

                // Table header
                table.Cell(1, 1).Range.Text = "№ п/п";
                table.Cell(1, 2).Range.Text = "Исходящий/входящий адрес электронной почты";
                table.Cell(1, 3).Range.Text = "Файл (КБ)";
                table.Cell(1, 4).Range.Text = "Категория срочности";
                table.Cell(1, 5).Range.Text = "Время приема/отправки";
                table.Cell(1, 6).Range.Text = "Кому (куда) адресована (адрес электронной почты)";
                table.Cell(1, 7).Range.Text = "Фамилия, инициалы и роспись дежурного по ШО";
                table.Cell(1, 8).Range.Text = "Примечание";
                table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Second row
                table.Cell(2, 1).Range.Text = "1";
                table.Cell(2, 2).Range.Text = "2";
                table.Cell(2, 3).Range.Text = "3";
                table.Cell(2, 4).Range.Text = "4";
                table.Cell(2, 5).Range.Text = "5";
                table.Cell(2, 6).Range.Text = "6";
                table.Cell(2, 7).Range.Text = "7";
                table.Cell(2, 8).Range.Text = "8";
                table.Rows[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                //
                // Third row in the end of method
                //

                // Fourth row
                int intRow = 4;

                // Retrieve the data and insert into new rows.
                object beforeRow = Type.Missing;

                foreach (var mail in emails)
                {
                    table.Rows.Add(ref beforeRow);

                    table.Cell(intRow, 1).Range.Text = counter++.ToString();
                    table.Cell(intRow, 2).Range.Text = mail.SenderAddress;
                    table.Cell(intRow, 3).Range.Text = mail.Attachments;
                    table.Cell(intRow, 4).Range.Text = mail.Category;
                    table.Cell(intRow, 5).Range.Text = mail.Date.ToShortTimeString();
                    table.Cell(intRow, 6).Range.Text = mail.RecivedAddress;
                    table.Cell(intRow, 7).Range.Text = " ";
                    table.Cell(intRow, 8).Range.Text = mail.Subject;

                    intRow += 1;
                }

                // Merge cells in third row and insert date from datepicker 
                table.Rows[3].Cells[7].Merge(table.Rows[3].Cells[8]);
                table.Rows[3].Cells[6].Merge(table.Rows[3].Cells[7]);
                table.Rows[3].Cells[5].Merge(table.Rows[3].Cells[6]);
                table.Rows[3].Cells[4].Merge(table.Rows[3].Cells[5]);
                table.Rows[3].Cells[3].Merge(table.Rows[3].Cells[4]);
                table.Rows[3].Cells[2].Merge(table.Rows[3].Cells[3]);
                table.Rows[3].Cells[1].Merge(table.Rows[3].Cells[2]);
                table.Cell(3, 1).Range.Text = selectedDate.ToLongDateString();
                table.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                winword.Visible = true;
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

        }
    }
}
