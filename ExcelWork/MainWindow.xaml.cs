using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Microsoft.Win32;

namespace ExcelWork
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private string file1 = "";
        private string file2 = "";
        private string file3 = "";

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();
            button.Background = Brushes.Red;
            button.Content = "No file!";
            button_Copy.Background = Brushes.Red;
            button_Copy.Content = "No file!";
            button_Copy1.Background = Brushes.Red;
            button_Copy1.Content = "No file!";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            file1 = of.FileName;
            button.Background = Brushes.Green;
        }

        private void Button_Copy_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            file2 = of.FileName;
            button_Copy.Background = Brushes.Green;
        }

        private void Button_Copy1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.ShowDialog();
            file3 = of.FileName;
            button_Copy1.Background = Brushes.Green;
        }
    }
}
