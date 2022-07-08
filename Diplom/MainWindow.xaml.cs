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
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace UniversalDocGenerator
{
    /// <summary>
    /// Логика взаимодействия для FSBWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void Window_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
  // path of Document              var helper = new WordHelper("C:/Users/user/source/repos/UniversalDocumentGenerator/Diplom/WordPattern/PatternDoc.doc");

                var items = new Dictionary<string, string>
                {
                    {
                        "<REQUEST_NUMBER>", TS_NumberRequest_TextBox.Text
                    },

                    {
                        "<TS_NAME>", TS_Name_TextBox.Text
                    },

                    {
                        "<TS_MANUFACTURER>", TS_Manufacturer_TextBox.Text
                    },

                    {
                        "<TS_MODEL>", TS_Model_TextBox.Text
                    },

                    {
                        "<TS_SERIALNUMBER>", TS_SerialNumber_TextBox.Text
                    },

                    

                };

                helper.Process(items);
            }

            catch
            {
                Console.WriteLine("Файл не найден");
            }
        }
    }
}
