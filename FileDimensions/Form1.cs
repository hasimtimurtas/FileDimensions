using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FileDimensions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Title = "Dosya Seçiniz";
            //openFileDialog1.InitialDirectory = "C:\\";
            //openFileDialog1.ShowDialog();
            //textBox1.Text = openFileDialog1.FileName;

            folderBrowserDialog1.ShowDialog();
            //textBox1.Text = folderBrowserDialog1.SelectedPath;

            CreateDimensionsFile(folderBrowserDialog1.SelectedPath);
        }

        private void CreateDimensionsFile(string path)
        {
            try
            {
                //Create the data set and table
                DataSet ds = new DataSet("New_DataSet");
                DataTable dt = new DataTable();
                dt.Clear();
                dt.Columns.Add("Filename");
                dt.Columns.Add("Width");
                dt.Columns.Add("Height");

                if (!File.Exists(path + "\\Kodlayici Kodlama Dosyasi.xlsx"))
                {
                    MessageBox.Show("Kodlayici Kodlama Dosyasi.xlsx file is not found!");
                    return;
                }
                
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(path + "\\Kodlayici Kodlama Dosyasi.xlsx", FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                    file.Close();
                }

                List<int> daylistindex = new List<int>();

                for (int i = 1; i <= 31; i++)
                {
                    daylistindex.Add(2);
                }


                string[] filenames = Directory.GetFiles(path, "*.jpg");
                foreach (var filename in filenames)
                {



                    System.Drawing.Image img = System.Drawing.Image.FromFile(filename);
                    //MessageBox.Show("Filename: " + filename + " Width: " + img.Width + ", Height: " + img.Height);
                    DataRow _ravi = dt.NewRow();
                    _ravi["Filename"] = filename;
                    _ravi["Width"] = img.Width;
                    _ravi["Height"] = img.Height;
                    dt.Rows.Add(_ravi);


                    string purefileName = filename.Substring(path.Length + 2);
                    int fileday = GetFileDay(purefileName);

                    XSSFSheet sheet = (XSSFSheet)hssfwb.GetSheet(fileday.ToString());

                    int index = daylistindex[fileday - 1];
                    var row = sheet.GetRow(index);

                    var celltemp = row.GetCell(0);

                    if (celltemp != null)
                    {
                        // TODO: you can add more cell types capatibility, e. g. formula
                        switch (celltemp.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                //MessageBox.Show(celltemp.NumericCellValue.ToString());
                                break;
                            case NPOI.SS.UserModel.CellType.String:
                                //MessageBox.Show(celltemp.StringCellValue);
                                break;
                        }
                    }

                    // ICreationHelper cH = hssfwb.GetCreationHelper();

                    var cell = row.CreateCell(7);
                    cell.SetCellValue(img.Width);
                    cell = row.CreateCell(8);
                    cell.SetCellValue(img.Height);                    

                    for (int i = 9; i < 23; i++)
                    {
                        cell = row.CreateCell(i);
                        cell.SetCellValue(0);
                    }

                    daylistindex[fileday - 1] = index + 1;
                }

                using (FileStream file = new FileStream(path + "\\Kodlayici Kodlama Dosyasi.xlsx", FileMode.Create, FileAccess.Write))
                {
                    hssfwb.Write(file);
                    file.Close();
                }

                //Add the table to the data set
                ds.Tables.Add(dt);

                //Here's the easy part. Create the Excel worksheet from the data set
                // ExcelLibrary.DataSetHelper.CreateWorkbook(path + "\\Dimensions.xls", ds);
                
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.Message);
            }

            MessageBox.Show(path + "Kodlayici Kodlama Dosyasi.xlsx is modified!!");
        }

        private int GetFileDay(string filename)
        {
            return Convert.ToInt32(filename.Substring(5, 2)); ;
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
    }
}
