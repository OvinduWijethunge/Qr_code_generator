using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelToQR
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            this.DataContext = this;
        }
        string workbookPath = "";
        string folderpath = "";
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            txtblk.Text = " ";
            txtblk2.Text = " ";

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();


            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Text documents (.xlsx)|*.xlsx";


            Nullable<bool> result = dlg.ShowDialog();


            if (result == true)
            {

                workbookPath = dlg.FileName;
                FileNameTextBox.Text = workbookPath;

            }

        }


        private void button2_Click(object sender, RoutedEventArgs e)
        {

            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();

            DialogResult result = dlg.ShowDialog();

            if (!string.IsNullOrWhiteSpace(dlg.SelectedPath))
            {

                folderpath = dlg.SelectedPath;
                FileNameTextBox2.Text = folderpath;

            }


        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                System.IO.FileStream fs = new
System.IO.FileStream(workbookPath, System.IO.FileMode.Open,System.IO.FileAccess.Read);
                fs.CopyTo(stream);


                using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    int rowCount = worksheet.Dimension.End.Row;

                    string newfolderPath = "";
                    var folderName = "";
                    for (int row = 3; row <= rowCount; row++)
                    {
                        var serialNum = "";
                        var issoftware = "";
                        var iswindows = "";
                        var ID = worksheet.Cells[row, 1].Value == null ?"No values" : worksheet.Cells[row, 1].Value.ToString();
                        var Type = worksheet.Cells[row, 2].Value == null ? "No values" : worksheet.Cells[row, 2].Value.ToString();
                        var status = worksheet.Cells[row, 3].Value == null ? "No values" : worksheet.Cells[row, 3].Value.ToString();
                        var user = worksheet.Cells[row, 4].Value == null ? "No values" : worksheet.Cells[row, 4].Value.ToString();

                        if (Type == "laptop")
                        {
                            if (worksheet.Cells[row, 5].Value == null)
                            {

                                serialNum = "Please Input values!!";
                            }
                            else
                            {
                                serialNum = worksheet.Cells[row,5].Value.ToString();
                            }
                            if (worksheet.Cells[row, 6].Value == null)
                            {

                                issoftware = "Please Input values!!";
                            }
                            else
                            {
                                issoftware = worksheet.Cells[row,6].Value.ToString();
                            }
                            if (worksheet.Cells[row, 7].Value == null)
                            {

                                iswindows = "Please Input values!!";
                            }
                            else
                            {
                                iswindows = worksheet.Cells[row,6].Value.ToString();
                            }

                        }
                        else
                        {
                            serialNum = worksheet.Cells[row, 5].Value ==null ? "No values" : worksheet.Cells[row, 5].Value.ToString();
                            issoftware = worksheet.Cells[row, 6].Value== null ? "No values" : worksheet.Cells[row, 6].Value.ToString();
                            iswindows = worksheet.Cells[row, 7].Value == null ? "No values" : worksheet.Cells[row, 7].Value.ToString();


                        }

                        var rowValue = "Inventory ID = " + ID + "\n" +"Inventory Type = " + Type + "\n" + "Rented  Cubo =  " + status + "\n" +"User = " + user + "\n" +
                                        "Serial Number = " + serialNum +"\n" + "Softwares Installed = " + issoftware + "\n" + "Windows Installed= " + iswindows;









                         var qrId = ID;


                        QRCodeGenerator qrGenerator = new QRCodeGenerator();
                        QRCodeData qrCodeData = qrGenerator.CreateQrCode(rowValue, QRCodeGenerator.ECCLevel.Q);
                        QRCode qrCode = new QRCode(qrCodeData);
                        Bitmap bmp = qrCode.GetGraphic(20, System.Drawing.Color.Black, System.Drawing.Color.White, true);

                        folderName = DateTime.Now.ToString("dddd - ddMMMM yyyy HH - mm");


                         newfolderPath = folderpath + "/" + folderName;
                        if (!Directory.Exists(newfolderPath))
                        {
                            Directory.CreateDirectory(newfolderPath);
                        }

                        using (MemoryStream ms = new MemoryStream())
                        {

                            bmp.Save(ms,System.Drawing.Imaging.ImageFormat.Png);
                            byte[] byteImage = ms.ToArray();
                            System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                            img.Save(newfolderPath + "/" + qrId +".Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);

                        }

                    }
                    FileNameTextBox.Text = "";
                    FileNameTextBox2.Text = "";
                    if (FileNameTextBox.Text != " " &&
FileNameTextBox2.Text != " ")
                    {
                        txtblk.Text = "File was successfully copied!!!!";
                         txtblk2.Text = newfolderPath;
                    }


                }
            }

        }

        private void button4_Click(object sender, RoutedEventArgs e)
        {
            FileNameTextBox.Text = "";
        }

        private void button5_Click(object sender, RoutedEventArgs e)
        {
            FileNameTextBox2.Text = "";
        }

        private void Grid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            txtblk.Text = " ";
            // txtblk2.Text = " ";
        }
    }
}