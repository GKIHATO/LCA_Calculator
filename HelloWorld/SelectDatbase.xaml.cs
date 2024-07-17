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

namespace HelloWorld
{
    /// <summary>
    /// Interaction logic forSelectDatabase.xaml
    /// </summary>
    public partial class SelectDatabase : Window
    {
        private Dictionary<string, List<String>> databaseList;

        public string databaseSelected;
        public SelectDatabase(Dictionary<string,List<String>> DatabaseList)
        {
            InitializeComponent();

            databaseList = DatabaseList;

            foreach (string item in databaseList.Keys)
            {
                CategoryComboBox.Items.Add(item);
            }
        }

        private void ok_btn_Click(object sender, RoutedEventArgs e)
        {
            if(NameComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please select a database");
                return;
            }

            DialogResult= true;
        }

        private void cancel_btn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void CategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoryComboBox.SelectedItem == null)
            {
                return;
            }

            NameComboBox.Items.Clear();

            foreach (string item in databaseList[CategoryComboBox.SelectedItem.ToString()])
            {
                NameComboBox.Items.Add(item);
            }
        }

        private void NameComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            databaseSelected = NameComboBox.SelectedItem.ToString();
        }
    }
}
