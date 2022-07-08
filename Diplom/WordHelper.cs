using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace UniversalDocGenerator
{
    class WordHelper
    {
        private FileInfo fileInfo;
       
        public WordHelper(string filename)
        {
            if(File.Exists(filename))
            {
                fileInfo = new FileInfo(filename);
            }
            else
            {
                throw new Exception("Файл не найден");
            }
        }

        internal void Process(Dictionary<string, string> items)
        {
            try
            {
                var app = new Word.Application();
                Object file = fileInfo.FullName;
                Object missing = Type.Missing;

                app.Documents.Open(file);

                foreach(var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

               
                    string path = @"\\Diplom\SaveDoc\";
                    string time = DateTime.Now.ToString("yyyymmDD");
                    string newFileName = Path.Combine(path, time);
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();
                

                
                MessageBox.Show("Документы сохранены");

                //return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //return false;



        }
    }
}
