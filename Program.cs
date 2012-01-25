using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace WordToMarkdown
{
    class Program
    {
        object missing = Missing.Value;
        object fullFilePath = Environment.CurrentDirectory + @"\TestDocument.docx";
        object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
        object noEncodingDialog = true; // http://msdn.microsoft.com/en-us/library/bb216319(office.12).aspx
        object f = false;
        object t = true;
        private static string pathToSublimeText = @"C:\Program Files\Sublime Text 2\sublime_text.exe";
        private static string outputFile = Environment.CurrentDirectory + @"\" + Guid.NewGuid().ToString() + ".md";

        static void Main(string[] args)
        {
            new Program();
            Console.ReadKey();
        }

        public Program()
        {
            Application word = LoadWordDocument();

            // convert tables to text
            for (int i = word.Selection.Document.Tables.Count; i > 0; i--)
            {
                word.Selection.Document.Tables[i].ConvertToText();
            }

            // replace headings
            ReplaceHeadings(word);

            // convert no number lists to indent
            ReplaceListNoNumber(word);

            // replace bold
            bool replaceOneBold = true;
            while (replaceOneBold)
            {
                replaceOneBold = ReplaceOneBold(word);
            }

            // replace italic
            bool replaceOneItalic = true;
            while (replaceOneItalic)
            {
                replaceOneItalic = ReplaceOneItalic(word);
            }

            // replace lists
            ReplaceLists(word);

            word.ActiveDocument.SaveAs2(outputFile, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSText);

            OpenSublimeText(outputFile);
        }

        static void OpenSublimeText(string f)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = pathToSublimeText;
            startInfo.Arguments = f;

            try
            {
                Process.Start(startInfo);
            }
            catch
            {
                startInfo.FileName = "Notepad.exe";
                startInfo.Arguments = f;
                Process.Start(startInfo);
            }
        }

        private static void ReplaceListNoNumber(Application word)
        {
            for (int i = word.Selection.Document.Paragraphs.Count; i > 0; i--)
            {
                Paragraph para = word.Selection.Document.Paragraphs[i];

                if (para.Range.ListFormat.ListType == WdListType.wdListNoNumbering)
                {
                    if (para.LeftIndent > 0)
                    {
                        para.Range.InsertBefore(">");
                    }
                    //waw needed? para.Range.InsertBefore(Environment.NewLine);
                }
            }
        }

        private void ReplaceHeadings(Application word)
        {
            for (int i = 1; i < 7; i++)
            {
                word.Selection.HomeKey(WdUnits.wdStory);

                bool replaceHeading = true;
                while (replaceHeading)
                {
                    replaceHeading = ReplaceHeading(word, i);
                }
            }
        }

        private Application LoadWordDocument()
        {
            object wordObject = null;
            Application word = null;

            try
            {
                wordObject = Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                // Do nothing.
            }

            if (wordObject != null)
            {
                word = (Application)wordObject;
            }
            else
            {
                word = new Application();
            }

            //word.Visible = true;

            word.Documents.Open(ref fullFilePath, ref t, ref f, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref encoding, ref missing, ref missing, ref missing, ref noEncodingDialog, ref missing);

            
            return word;
        }

        private bool ReplaceHeading(Application word, int number)
        {
            bool replaceHeading = false;
            string replacement = string.Empty;

            switch (number)
            {
                case 1:
                    replacement = "#";
                    break;
                case 2:
                    replacement = "##";
                    break;
                case 3:
                    replacement = "###";
                    break;
                case 4:
                    replacement = "####";
                    break;
                case 5:
                    replacement = "#####";
                    break;
                case 6:
                    replacement = "######";
                    break;
            }

            object heading = word.ActiveDocument.Styles["Heading " + number];
            object normal = word.ActiveDocument.Styles["Normal"];

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.set_Style(ref heading);

            while (word.Selection.Find.Execute())
            {
                replaceHeading = true;
                word.Selection.Range.InsertBefore(replacement + " ");
                word.Selection.set_Style(ref normal);
                word.Selection.Find.Execute();
            }

            return replaceHeading;
        }

        private bool ReplaceOneBold(Application word)
        {
            object findText = "";
            bool replaceOneBold = false;

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.Font.Bold = 1;
            word.Selection.HomeKey(WdUnits.wdStory);

            while (word.Selection.Find.Execute())
            {
                replaceOneBold = true;
                word.Selection.Text = "**" + word.Selection.Text + "**";
                word.Selection.Font.Bold = 0;
                word.Selection.Find.Execute();
            }

            return replaceOneBold;
        }

        private bool ReplaceOneItalic(Application word)
        {
            object findText = "";
            bool replaceOneItalic = false;

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.Font.Italic = 1;
            word.Selection.HomeKey(WdUnits.wdStory);

            while (word.Selection.Find.Execute())
            {
                replaceOneItalic = true;
                word.Selection.Text = "_" + word.Selection.Text + "_";
                word.Selection.Font.Italic = 0;
                word.Selection.Find.Execute();
            }

            return replaceOneItalic;
        }

        private void ReplaceLists(Application word)
        {
            word.Selection.HomeKey(WdUnits.wdStory);

            for (int i = word.Selection.Document.Paragraphs.Count; i > 0; i--)
            {
                try
                {
                    for (int j = word.Selection.Document.Lists[i].ListParagraphs.Count; j > 0; j--)
                    {
                        Paragraph para = word.Selection.Document.Lists[i].ListParagraphs[j];

                        if (para.Range.ListFormat.ListType == WdListType.wdListBullet)
                        {
                            para.Range.InsertBefore(ListIndent(para.Range.ListFormat.ListLevelNumber, "*"));
                        }

                        if (para.Range.ListFormat.ListType == WdListType.wdListSimpleNumbering ||
                            para.Range.ListFormat.ListType == WdListType.wdListMixedNumbering ||
                            para.Range.ListFormat.ListType == WdListType.wdListListNumOnly)
                        {
                            para.Range.InsertBefore(para.Range.ListFormat.ListValue + ". ");
                        }
                    }

                    word.Selection.Document.Lists[i].Range.InsertParagraphBefore();
                    word.Selection.Document.Lists[i].Range.InsertParagraphAfter();
                    word.Selection.Document.Lists[i].RemoveNumbers();
                }
                catch
                { }
            }
        }

        private string ListIndent(int number, string text)
        {
            string returnValue = "";

            for (int i = 0; i < number -1; i++)
            {
                returnValue = returnValue + "    ";
            }

            returnValue = returnValue + text + "    ";

            return returnValue;
        }

    }


}
