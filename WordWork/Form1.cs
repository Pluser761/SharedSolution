using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CheckBox = System.Windows.Forms.CheckBox;

namespace WordWork
{
    public partial class Form1 : Form
    {
        Word._Application oWord;
        _Document oDoc;
        string name;
        int labels = 0;
        int texboxes = 0;

        public Form1(string file_name, string path)
        {
            this.oWord = new Word.Application();
            this.oDoc = oWord.Documents.Add(path);
            this.name = file_name;
            InitializeComponent();
        }

        ~Form1()
        {
            oWord.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int i = 0;
            foreach (Bookmark bookmark in oDoc.Bookmarks)
            {
                Controls.Add(create_lab(i, bookmark.Name));
                Controls.Add(create_box(i));
                Controls.Add(create_check(i));
                i++;
            }
            Controls.Add(create_but(i));
            Height = 10 + (i + 2) * 31 + 10;
            Width = 645 + 31 + 30;
        }

        private Label create_lab(int num, string text)
        {
            Label label = new Label();
            label.Top = 10 + num * 31;
            label.Left = 10;
            label.Width = 130;
            label.Height = 26;
            label.Text = text;
            label.Name = "Label" + labels;
            labels++;
            return label;
        }

        private TextBox create_box(int num)
        {
            TextBox textBox = new TextBox();
            textBox.Top = 13 + num * 31;
            textBox.Left = 145;
            textBox.Width = 500;
            textBox.Height = 23;
            textBox.Text = "";
            textBox.Name = "TextBox" + texboxes;
            texboxes++;
            return textBox;
        }

        private Button create_but(int num)
        {
            Button button = new Button();
            button.Top = 10 + num * 31;
            button.Left = 515;
            button.Width = 130;
            button.Height = 26;
            button.Text = "Выполнить";
            button.Name = "Button";
            button.Click += new EventHandler(method);
            return button;
        }

        private CheckBox create_check(int num)
        {
            CheckBox check = new CheckBox();
            check.Top = 10 + num * 31;
            check.Left = 515 + 130 + 15;
            check.Name = "Check" + num;
            return check;
        }

        private void method(object sender, EventArgs e)
        {
            int i = 0;
            foreach (Bookmark bookmark in oDoc.Bookmarks)
            {
                //bookmark.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                string text = Controls["TextBox" + i].Text;
                CheckBox check = Controls["Check" + i] as CheckBox;
                if (check.Checked)
                {
                    String str = RusNumber.Str(int.Parse(text)).TrimEnd(' ');
                    text = text + " (" + str + ")";
                }
                bookmark.Range.Text = text;
                i++;
            }
            oDoc.SaveAs(Environment.CurrentDirectory + "\\" + name + ".docx");
            oDoc.Close();
            Close();
        }
        
    }
}
