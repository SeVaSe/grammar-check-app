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
using Microsoft.Office.Interop.Word;


namespace pr4
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Add();
            Range range = doc.Range();

            range.Text = Txt_box.Text;
            int err = doc.SpellingErrors.Count;

            if (range.SpellingErrors.Count > 0)
            {
                string msg = "Орфографические ошибки: \n";
                foreach (Range error in range.SpellingErrors)
                {
                    msg += error.Text + "\n";
                }
                MessageBox.Show(msg, $"Найдены ошибки в орфографии, количество: {err}");
            }
            else
            {
                MessageBox.Show("Ошибок в тексте нет", "Корректный синтаксис");
            }
        }
    }
}
