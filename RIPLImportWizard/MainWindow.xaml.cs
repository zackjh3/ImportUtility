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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Data;

namespace RIPLImportWizard
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MappingWindow mappingWin = new MappingWindow();
        public MainWindow()
        {
            InitializeComponent();
        }
        public static OpenFileDialog openfile = new OpenFileDialog();

        private void btn_FileSearch(object sender, RoutedEventArgs e)
        {
           
            int selection = SourceDatabaseType.SelectedIndex;
            openfile.InitialDirectory = "C:\\Users\\zach.hine\\American Innovations\\Import";
            if (selection == 0)
            {
                openfile.DefaultExt = ".mdb";
                openfile.Filter = "(.mdb)|*.mdb";
            }
            else if (selection == 1)
            {
                openfile.DefaultExt = ".accdb";
                openfile.Filter = "(.accdb)|*.accdb";
            }
            else if (selection==2)
            {
                openfile.DefaultExt = ".xls";
                openfile.Filter = "(.xls)|*.xls";
            }
            else if (selection==3)
            {
                openfile.DefaultExt = ".xlsx";
                openfile.Filter = "(.xlsx)|*.xlsx";
            }
            var browsefile = openfile.ShowDialog();
            if (browsefile == true)
            {
                filePath.Text = openfile.FileName;
                Start.IsEnabled = true;
            }
        }
        public void Button_Click(object sender, RoutedEventArgs e)
        {
            // Navigate to mapping
            
            this.Hide();
            mappingWin.ShowDialog();
    
        }
    }
}
