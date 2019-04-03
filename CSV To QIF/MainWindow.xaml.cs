using System;
using System.Collections.Generic;
using System.Globalization;
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
using CsvHelper;

namespace CSV_To_QIF
{

    public class CSVLine
    {
        public string Date { get; set; }
        public string Description { get; set; }
        public string Amount { get; set; }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string fileNameString = "";

        private void openFile(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".csv";
            dlg.Filter = "CSV Files (*.csv)|*.csv";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                fileNameString = dlg.FileName;
                FileName.Content = "Filename:"+dlg.FileName;
            }
        }

        private void SaveFile(object sender, RoutedEventArgs e)
        {
            if (fileNameString == "" || fileNameString == null)
            {
                MessageBox.Show("Please select a CSV file first.");
                return;
            }

            try
            {
                var s = new System.IO.StreamReader(fileNameString);
                using (var reader = new CsvHelper.CsvReader(s,new CsvHelper.Configuration.Configuration() {  }))
                {
                    var records = reader.GetRecords<CSVLine>();
                    List<string> quickenFileContent = new List<string>() { "!Type:CCard" };
                    foreach (var record in records)
                    {
                        var date = DateTime.ParseExact(record.Date, InDate.Text, CultureInfo.InvariantCulture);
                        quickenFileContent.Add("D" + date.ToString(OutDate.Text) + "'" + date.ToString("yyyy"));
                        quickenFileContent.Add("N");
                        quickenFileContent.Add("T" + record.Amount);
                        quickenFileContent.Add("P" + record.Description);
                        quickenFileContent.Add("^");
                    }
                    var saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                    saveFileDialog.Filter = "Quicken Import File (*.QIF)|*.qif";
                    saveFileDialog.Title = "Save Quicken file";
                    saveFileDialog.DefaultExt = "*.qif";
                    saveFileDialog.FileName = "New Import.qif";
                    saveFileDialog.ShowDialog();

                    if (saveFileDialog.FileName == null)
                    {
                        MessageBox.Show("Did not save file.");
                        return;
                    }else if (saveFileDialog.FileName == "")
                    {
                        MessageBox.Show("Did not save file.");
                        return;
                    }
                    else
                    {
                        try
                        {
                            System.IO.File.WriteAllLines(saveFileDialog.FileName, quickenFileContent.ToArray());
                            MessageBox.Show("File successfully saved.");
                        }
                        catch (Exception f)
                        {
                            MessageBox.Show("Could not save file: " + f.Message);
                        }
                        
                    }

                }
                
            }
            catch (System.IO.FileNotFoundException)
            {
                MessageBox.Show("Couldn't open CSV file.");
            }
            
        }
    }
}
