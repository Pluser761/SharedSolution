using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using ExelWork;

namespace ExelWork
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            textBox.Text = of.FileName;
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            textBox1.Text = of.FileName;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            textBox2.Text = of.FileName;
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            textBox3.Text = of.FileName;
        }

        private void Button4_Click(object sender, RoutedEventArgs e)
        {
            Excel mainFile = new Excel(textBox.Text, textBoxSheet.Text);

            Excel[] arr = 
            {
                new Excel(textBox1.Text, textBoxSheet.Text),
                new Excel(textBox2.Text, textBoxSheet.Text),
                new Excel(textBox3.Text, textBoxSheet.Text)
            };

            int col1 = mainFile.columnToInt(textBox11.Text);

            for (int f = 0; f < 3; f++)
            {
                for (int i = mainFile.columnToInt(textBox11.Text); i <= mainFile.columnToInt(textBox21.Text); i++)
                    for (int k = Convert.ToInt32(textBox12.Text); k <= Convert.ToInt32(textBox22.Text); k++)
                    {
                        string temp = mainFile.ReadCell(k, i);
                        if (temp == "" || temp == null)
                            mainFile.WriteCell(arr[f].ReadCell(k, i), k, i);
                    }
                arr[f].Close();
            }

            

            mainFile.Save();
            mainFile.Close();

            System.Windows.Forms.MessageBox.Show("Готово!");
        }
        
    }
}