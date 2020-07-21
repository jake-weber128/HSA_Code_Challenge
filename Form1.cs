﻿using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace Jake_Test2
{
    public partial class Form1 : Form
    {
        public DataTable dt = new DataTable();
        
        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            //this.DragDrop += new DragEventHandler(this.Form1_DragDrop);
            //this.DragEnter += new DragEventHandler(this.Form1_DragEnter);

        }


        private void Form1_Load(object sender, EventArgs e)
        {
                   }
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
            {
                String[] strGetFormats = e.Data.GetFormats();
                e.Effect = DragDropEffects.None;
            }
        }
        #region GetFile
        private void tb_path_DragDrop(object sender, DragEventArgs e)
        {
            // Get the files being dragged and process them.

            string[] fileNames = null;
            TextBox PB = (TextBox)sender;
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
                {
                    fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);

                    foreach (string fileName in fileNames)
                        // File name action
                        // ProcessFile(fileName.ToString(), PB.Name);
                        PB.Text = fileName.ToString();
                }
                else if (e.Data.GetDataPresent("FileGroupDescriptor"))
                {
                    Stream theStream = (Stream)e.Data.GetData("FileGroupDescriptor");
                    byte[] fileGroupDescriptor = new byte[512];
                    theStream.Read(fileGroupDescriptor, 0, 512);
                    StringBuilder fileName = new StringBuilder("");
                    int i = 76;

                    while (fileGroupDescriptor[i] != 0)
                    {
                        fileName.Append(Convert.ToChar(fileGroupDescriptor[i]));
                        i += 1;
                    }

                    theStream.Close();
                    string path = @"C:\Users\" + CommonFunctions.GetUserName() + @"\AppData\";
                    string theFile = path + fileName.ToString();
                    MemoryStream ms = (MemoryStream)e.Data.GetData("FileContents", true);
                    byte[] fileBytes = new byte[ms.Length - 1 + 1];
                    ms.Position = 0;
                    ms.Read(fileBytes, 0, System.Convert.ToInt32(ms.Length));
                    FileStream fs = new FileStream(theFile, FileMode.Create);
                    fs.Write(fileBytes, 0, System.Convert.ToInt32(fileBytes.Length));
                    fs.Close();
                    FileInfo tempFile = new FileInfo(theFile);
                    PB.Text = tempFile.FullName.ToString();
                    //ProcessFile(tempFile.FullName.ToString(), PB.Name);
                }

            }
            // TreeView1.Nodes.Add(node)
            catch (Exception ex)
            {
                Trace.WriteLine("Error in DragDrop function: " + ex.Message);
            }
        }

        public void tb_path_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else if (e.Data.GetDataPresent("FileGroupDescriptor"))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"\\C:";

            OpenFileDialog result = new OpenFileDialog();
            // Show the dialog.
            result.InitialDirectory = path;
            // result.Multiselect = True
            if (result.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i <= Convert.ToInt32(result.FileNames.Count()); i++)
                    path = result.FileNames.GetValue(i).ToString();
            }
            tb_path.Text = path;

        }
        #endregion
        #region Validation
        private void rule3Check()
        {
            bool valid = true;
            string status = "Accepted";
            foreach (DataRow row in dt.Rows)
            {

                valid = Effectivedatecheck(Convert.ToDateTime(row.Field<DateTime>("EffectiveDate")));
                if (valid == false)
                {
                    row[5] = "Rejected";
                }
                else
                {
                    row[5] = "Accepted";
                }
                valid = AgeCheck(Convert.ToDateTime(row.Field<DateTime>("DOB")));

                if (valid == false)
                {
                    row[5] = "Rejected";
                }

            }
           
                           
        }
        private bool Rule2check()
        {
            bool valid1 = true;
            bool valid = true;

            foreach (DataRow row in dt.Rows)
            {
               
                  valid =  RowValidation(row);
                if (valid == false)
                {
                    MessageBox.Show("File Processing Stopped Due to Blank Record");
                    return false;
                }
                
                valid = ValidDateCheck(Convert.ToDateTime(row.Field<DateTime>("EffectiveDate")));
                if (valid == false)
                {
                    MessageBox.Show("File Processing Stopped Due to Invalid Date");
                    return false;
                }
                
                valid = ValidDateCheck(Convert.ToDateTime(row.Field<DateTime>("DOB")));
                if (valid == false)
                {
                    MessageBox.Show("File Processing Stopped Due to Invalid Date");
                    return false;
                }

                 valid=  PlanTypeCheck(row.Field<string>("PlanType"));
                if (valid == false)
                {
                    MessageBox.Show("File Processing Stopped Due to Bad Plan Type");
                    return false;
                }
            }

            return valid; 
        }
        private bool Effectivedatecheck(DateTime  effdate)
        {
            bool valid = false;
            DateTime currDate = DateTime.Now;
            TimeSpan span = currDate - effdate;
            int days = span.Days; 

            if(days <= 30 )
            {
                valid = true; 
            }

            return valid; 
        }
       private bool AgeCheck(DateTime DOB)
        {
            bool valid = false;

            DateTime zeroTime = new DateTime(1, 1, 1);

            DateTime a = DateTime.Now;
            DateTime b = Convert.ToDateTime(DOB);

            TimeSpan span =  a-b;
            // Because we start at year 1 for the Gregorian
            // calendar, we must subtract a year here.
            int years = (zeroTime + span).Year - 1;
            if (years >= 18 )
            {
                valid = true;
            }

            return valid; 

        }
        private bool RowValidation(DataRow row)
        {
            bool RowPassFail = true;
            for(int i = 0;i<=row.ItemArray.Length ; i++)
            {
                bool PassFail = true;
               PassFail = RequiredFieldCheck(row[i].ToString());
                if (PassFail == false)
                {
                    RowPassFail = false;
                }
            }
            return RowPassFail;
        }
        private bool RequiredFieldCheck(string value)
        {
            bool valid = false;

            if (value != "")
            {
                valid = true;
            }

            return valid;
        }
        private bool ValidDateCheck(DateTime value)
        {
            bool valid = false;
            DateTime results;

            if (DateTime.TryParseExact(value.ToString(), "MM/dd/yyyy", null, DateTimeStyles.None, out results) == true)
            {
                valid = true; 
            }

            return valid;
        }
        private bool PlanTypeCheck(string value)
        {
            bool valid = false;

            if (value.ToString() == "HSA"||value.ToString() == "HRA"|| value.ToString() =="FSA")
            {
                valid = true; 
            }

            return valid;
        }
        #endregion
        #region Processes
        private void PrintRowLine(DataRow row)
        {
            listBox1.Items.Add(row[5].ToString() + ", " + row[0].ToString() + ", " + row[1].ToString() + ", " + row.Field<DateTime>("DOB").ToShortDateString() + ", " + row[3].ToString() + ", " + row.Field<DateTime>("EffectiveDate").ToShortDateString());

        }
        public void ProcessFile(string path)
        {

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +";Extended Properties=\"Text;HDR=No\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
               // DataTable dataTable = new DataTable();
                dt.Locale = CultureInfo.CurrentCulture;
                dt.Columns.Add("FirstName");
                dt.Columns.Add("LastName");
                dt.Columns.Add("DOB");
                dt.Columns.Add("PlanType");
                dt.Columns.Add("EffectiveDate");
                dt.Columns.Add("Status");
                adapter.Fill(dt);
                return ;
            }
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            bool rule2pass = true;
            string rule3pass = ""; 

            ProcessFile(tb_path.ToString());

            rule2pass = Rule2check();
            if (rule2pass == false)
                return;
             rule3Check();

                foreach(DataRow row in dt.Rows)
            {
                PrintRowLine(row);
            }

            dt.Clear();
            MessageBox.Show("File Data Proccessed Successfully");

        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void Form1_DragEnter_1(object sender, DragEventArgs e)
        {

        }

        private void tb_path_DragDrop_1(object sender, DragEventArgs e)
        {

        }
    }

    class CommonFunctions
        {
            #region CommonFunctions
            public static string GetUserName()
            {
                if (System.Security.Principal.WindowsIdentity.GetCurrent() != null)
                {
                    string[] parts = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Split('\\');
                    string username = parts[1].ToString();
                    return username;
                }
                else
                {
                    return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                }
            }
            #endregion
        }


    }

