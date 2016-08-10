using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DocumentTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DsoOffice dsoOffice = null;

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void LoadDSO()
        {
            WindowsFormsHost host = new WindowsFormsHost();

            System.Windows.Forms.Panel panel = new System.Windows.Forms.Panel();
            dsoOffice = new DsoOffice();
            dsoOffice.Dock = System.Windows.Forms.DockStyle.Fill;
            panel.Controls.Add(dsoOffice);

            host.Child = panel;
            gridOffice.Children.Add(host);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDSO();
        }

        private void OpenWord_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "*.doc|*.doc|*.docx|*.docx";
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var filename = ofd.FileName;
                if( dsoOffice ==null)
                {
                    MessageBox.Show("控件异常");
                    return;
                }
                dsoOffice.OpenDocument(filename);
            }
        }

        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "*.xls|*.xls|*.xlsx|*.xlsx";
            ofd.Multiselect = false;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var filename = ofd.FileName;
                dsoOffice.OpenDocument(filename);
            }
        }

        private void SaveWord_Click(object sender, RoutedEventArgs e)
        {
            dsoOffice.SaveDocument();
        }

        private void SaveExcel_Click(object sender, RoutedEventArgs e)
        {
            dsoOffice.SaveDocument();
        }
    }
}
