using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


// Для офиса
using Microsoft.Office.Tools.Word;


using Word = Microsoft.Office.Interop.Word;


using System.Text.RegularExpressions;




namespace DipGen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        Word.Application WordApp;
        Word.Document DocRez;

        public MainWindow()
        {
            InitializeComponent();
        }


        private string FindRez(string student)
        {
            DocRez.Activate();
            string stud_rez = "";

            for (int cr = 0; cr < DocRez.Tables[2].Rows.Count; cr++)
            {
                DocRez.Tables[2].Cell(cr + 2, 4).Select();
                string str_stud = WordApp.Selection.Text;

                
                Match m = Regex.Match(str_stud, "\\S*" + student + "\\S*");
                int mnum = m.Length;
                if (m.Length > 0)
                {
                    DocRez.Tables[2].Cell(cr + 2, 2).Select();
                    stud_rez = WordApp.Selection.Text;

                    DocRez.Tables[2].Cell(cr + 2, 3).Select();
                    stud_rez = stud_rez + ", " + WordApp.Selection.Text;
                    stud_rez = stud_rez.Replace("\r", " ");
                    stud_rez = stud_rez.Replace("\a", "");
                    stud_rez = stud_rez.Replace("\v", "");

                    break;
                }

                
            }

            return stud_rez;
        
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            string DocName = @"d:\Denis\Дипломники\Приказ 2014 для генератора.doc";
            string PatternName = @"d:\Denis\Дипломники\Карпиевич.docx";
            string RezName = @"d:\Denis\Дипломники\Распоряжение о рецензентах.doc";

            
            // Создаем объект Word - равносильно запуску Word
            WordApp = new Word.Application();

            try
            {
                //WordApp.Documents.Close();

                // Делаем его видимым
                WordApp.Visible = true;

                // Открыть документ
                Word.Document DocMain = WordApp.Documents.Open(DocName);
                Word.Document DocPattern;
                DocRez = WordApp.Documents.Open(RezName);

                DocMain.Activate();
                int count = DocMain.Tables[1].Rows.Count - 2;

                for (int ci = 0; ci < count; ci++)
                {
                    


                    DocPattern = WordApp.Documents.Open(PatternName);
                                        
                    DocMain.Activate();

                    int tnum = DocMain.Tables.Count;

                    DocMain.Tables[1].Cell(ci + 3, 3).Select();
                    string str_fio = WordApp.Selection.Text;
                    str_fio = str_fio.Replace("\r", " ");
                    str_fio = str_fio.Replace("\a", "");
                    str_fio = str_fio.Replace("\v", "");

                    DocMain.Tables[1].Cell(ci + 3, 4).Select();
                    string str_pris = WordApp.Selection.Text;
                    str_pris = str_pris.Replace("\r", " ");
                    str_pris = str_pris.Replace("\a", "");
                    str_pris = str_pris.Replace("\v", "");

                    DocMain.Tables[1].Cell(ci + 3, 5).Select();
                    string str_tema = WordApp.Selection.Text;
                    str_tema = str_tema.Replace("\r", " ");
                    str_tema = str_tema.Replace("\a", "");
                    str_tema = str_tema.Replace("\v", "");
                    str_tema = str_tema.Remove(str_tema.Length - 1);
                    str_tema = "«" + str_tema + "»";

                    DocMain.Tables[1].Cell(ci + 3, 6).Select();
                    string str_ruk = WordApp.Selection.Text;
                    str_ruk = str_ruk.Replace("\r", " ");
                    str_ruk = str_ruk.Replace("\a", "");
                    str_ruk = str_ruk.Replace("\v", "");

                    DocMain.Tables[1].Cell(ci + 3, 2).Select();
                    string stud_rez = WordApp.Selection.Words[1].Text;
                    stud_rez = stud_rez.Replace("\r", " ");
                    stud_rez = stud_rez.Replace("\a", "");
                    stud_rez = stud_rez.Replace("\v", "");



                    //str_stud

                    //str_stud = str_stud.Replace("\r", " ");
                    //str_stud = str_stud.Replace("\a", "");
                    //str_stud = str_stud.Replace("\v", "");




                    string str_rez = FindRez(stud_rez);


                    DocPattern.Activate();

                    Word.Bookmark fio = DocPattern.Bookmarks["fio"];
                    Word.Bookmark tema = DocPattern.Bookmarks["tema"];
                    Word.Bookmark pris = DocPattern.Bookmarks["pris"];
                    Word.Bookmark ruk = DocPattern.Bookmarks["ruk"];
                    Word.Bookmark rez = DocPattern.Bookmarks["rez"];

                    fio.Range.Collapse();
                    WordApp.Selection.SetRange(fio.Range.Start, fio.Range.End);
                    WordApp.Selection.Text = str_fio;

                    tema.Range.Collapse();
                    WordApp.Selection.SetRange(tema.Range.Start, tema.Range.End);
                    WordApp.Selection.Text = str_tema;

                    pris.Range.Collapse();
                    WordApp.Selection.SetRange(pris.Range.Start, pris.Range.End);
                    WordApp.Selection.Text = str_pris;

                    ruk.Range.Collapse();
                    WordApp.Selection.SetRange(ruk.Range.Start, ruk.Range.End);
                    WordApp.Selection.Text = str_ruk;

                    rez.Range.Collapse();
                    WordApp.Selection.SetRange(rez.Range.Start, rez.Range.End);
                    WordApp.Selection.Text = str_rez;



                    string filename = "c:\\Протокол " + str_fio + ".docx";
                    DocPattern.SaveAs(filename); // Сохранить под другим именем
                    DocPattern.Close(Word.WdSaveOptions.wdSaveChanges);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                WordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); // Выход
            }

            WordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); // Выход





        }
    }
}
