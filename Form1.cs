using OfficeOpenXml;
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

namespace ExcelTool_ModifyExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                modifyFile(file);
            }
        }





        public void modifyFile(string filePath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //assume already got file path
            //string filePath = @"C:\Users\tcy70\source\repos\ModifyExcel\file999.xlsx";

            // open excel
            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                // check 
                if (package.Workbook.Worksheets.Count > 0)
                {
                    //set backup
                    string currentDirectory = Directory.GetCurrentDirectory();
                    // get file name
                    string fname = Path.GetFileName(filePath);
                    //string backupFname = fname.Replace(".", "(Backup).");
                    string backupFname = GenerateUniqueBackupName(fname, currentDirectory);

                    // get current path


                    // struct the path
                    string backupDirectory = Path.Combine(currentDirectory, "backup");
                    string relativePath = Path.Combine(backupDirectory, backupFname);

                    // if doesn't exist then create
                    if (!Directory.Exists(backupDirectory))
                    {
                        Directory.CreateDirectory(backupDirectory);
                    }

                    // save backup
                    package.SaveAs(new FileInfo(relativePath));

                    Console.WriteLine("Backup sucess");
                    Console.ReadLine();

                    // get first one worksheets
                    var worksheet = package.Workbook.Worksheets[0];

                    // get range(start and end)
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    // create a uniquepair to store group of a value & e value that have appeare 
                    var uniquePairs = new HashSet<(string, string)>();

                    // check each row's A & E column value
                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        string aValue = worksheet.Cells[row, 1].Text; // value A column
                        string eValue = worksheet.Cells[row, 5].Text; // value e column

                        // if A, E column value exist
                        if (!string.IsNullOrEmpty(aValue) && !string.IsNullOrEmpty(eValue))
                        {

                            // pair for A & E column value
                            var pair = (aValue, eValue);
                            //if (aValue != eValue)//setting baru
                            //{
                            // // if already exist, mean duplicate, so delete current row
                            if (uniquePairs.Contains(pair))
                            {
                                worksheet.DeleteRow(row);
                                row--; //After deleting a row, reduce the row index by 1.
                            }
                            else
                            {
                                // add to uniquepair
                                uniquePairs.Add(pair);

                                //check if the main part and sub part are same value
                                if (aValue == eValue)
                                {
                                    //if same then put blank
                                    worksheet.Cells[row, 5].Value = ""; // put blank
                                    worksheet.Cells[row, 8].Value = "";
                                }

                            }
                            //}
                            //else//setting baru
                            //{
                            //worksheet.DeleteRow(row);
                            //row--; //After deleting a row, reduce the row index by 1.
                            //}
                        }
                    }


                    ////////////////////////////////////////////////////////////////////////////////////////////
                    ///
                    // // create a uniquepair to store group of a value & C value that have appeare 
                    var uniquePairs2 = new HashSet<(string, string)>();

                    // check each row's A & C column value
                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        string aValue = worksheet.Cells[row, 1].Text; // A 
                        string cValue = worksheet.Cells[row, 3].Text; // C
                        string paValue;
                        string pcValue;
                        if (row > 1)
                        {
                            paValue = worksheet.Cells[row - 1, 1].Text;
                            pcValue = worksheet.Cells[row - 1, 3].Text;
                        }
                        else
                        {
                            paValue = "Null";
                            pcValue = "Null";
                        }

                        // if A, C column value exist
                        if (!string.IsNullOrEmpty(aValue) && !string.IsNullOrEmpty(cValue))
                        {
                            // pair for A & C column value
                            var pair = (aValue, cValue);

                            // if already exist, mean duplicate, so delete current row's C column value
                            if (uniquePairs2.Contains(pair))
                            {
                                if (paValue == aValue && pcValue == cValue || paValue == aValue && pcValue == "")
                                {
                                    worksheet.Cells[row, 3].Value = ""; // put blank
                                }
                            }
                            else
                            {
                                // add to uniquepair
                                uniquePairs2.Add(pair);
                            }
                        }
                    }


                    ////////////////////////////////////////////////////////////////////////////////////////////
                    ///
                    // // create a uniquepair to store group of a value & G value that have appeare 
                    var uniquePairs3 = new HashSet<(string, string)>();

                    // check each row's A & G column value
                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        string aValue = worksheet.Cells[row, 1].Text; // A column value
                        string gValue = worksheet.Cells[row, 7].Text; // G column value
                        string paValue;
                        string pgValue;
                        if (row > 1)
                        {
                            paValue = worksheet.Cells[row - 1, 1].Text;
                            pgValue = worksheet.Cells[row - 1, 7].Text;
                        }
                        else
                        {
                            paValue = "Null";
                            pgValue = "Null";
                        }

                        // if a,g column value exist
                        if (!string.IsNullOrEmpty(aValue) && !string.IsNullOrEmpty(gValue))
                        {
                            // pair for A & G column value
                            var pair = (aValue, gValue);

                            // if already exist, mean duplicate, so delete current row's G column value
                            if (uniquePairs3.Contains(pair))
                            {
                                if (paValue == aValue && pgValue == gValue || paValue == aValue && pgValue == "")
                                {
                                    worksheet.Cells[row, 7].Value = ""; // put blank
                                }
                            }
                            else
                            {
                                // add to uniquepair
                                uniquePairs3.Add(pair);
                            }
                        }
                    }

                    // save excel
                    //package.Save();
                    //string backupFilePath= filePath.Replace(".", "(Backup).");
                    package.SaveAs(new System.IO.FileInfo(@filePath));
                    //Console.WriteLine(filePath);
                    MessageBox.Show("File has been modified and store original backup in " + relativePath + ".");
                    //MessageBox.Show("File path:"+filePath);
                }
                else
                {
                    MessageBox.Show("The File Don't have any Worksheets.");

                }

            }


        }

        static string GenerateUniqueBackupName(string originalName, string currentDirectory)
        {
            // struct backup filename
            string backupName = originalName;
            int counter = 1;

            // struct path
            string backupDirectory = Path.Combine(currentDirectory, "backup");

            // check how many same file exist
            while (File.Exists(Path.Combine(backupDirectory, backupName)))
            {
                backupName = $"{Path.GetFileNameWithoutExtension(originalName)}({counter}){Path.GetExtension(originalName)}";
                counter++;
            }

            return backupName;
        }
    }
}
