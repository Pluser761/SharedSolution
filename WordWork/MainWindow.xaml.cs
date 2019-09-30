using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
namespace WordWork
{


    public partial class MainWindow : System.Windows.Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Copy_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog of = new System.Windows.Forms.OpenFileDialog();
            of.ShowDialog();
            textBox_Copy4.Text = of.FileName;
        }
        /*
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Word._Application oWord = new Word.Application();
            _Document oDoc = oWord.Documents.Add(textBox_Copy4.Text);
            //заполнение закладок по именам в word

            oDoc.Bookmarks["NAME"].Range.Text = textBox.Text;
            oDoc.Bookmarks["NAME1"].Range.Text = textBox.Text;
            oDoc.Bookmarks["SECONDNAME"].Range.Text = textBox_Copy.Text;
            oDoc.Bookmarks["DAY"].Range.Text = textBox_Copy1.Text;
            oDoc.Bookmarks["MONTH"].Range.Text = textBox_Copy2.Text;
            oDoc.Bookmarks["YEAR"].Range.Text = textBox_Copy3.Text;
            //сохранение документа
            oDoc.SaveAs(Environment.CurrentDirectory + "\\" + textBox.Text + ".docx");
            oDoc.Close();

            System.Windows.Forms.MessageBox.Show("Готово!");
        }
        */

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            Form1 f = new Form1(textBox.Text, textBox_Copy4.Text);
            f.Show();
        }
        
    }
}
