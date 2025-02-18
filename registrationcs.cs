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

namespace CARDIOLOGY
{
    public partial class registrationcs : Form
    {
        private string uploadedFilePath; // Class-level variable to store the uploaded file path
        private ExcelPackage package;
        public static int g_rHandle, g_retCode;
        public static bool g_isConnected = false;
        public static byte g_Sec;
        public static byte[] g_pKey = new byte[6];
        ComboBox cbPort = new ComboBox();
        int iTips;
        int iAttempt;
        string CardBalance, dstrpass1;

        private void registrationcs_Load(object sender, EventArgs e)
        {
            Connect();
            cbPort.Items.Add("USB1");
            cbPort.SelectedIndex = 0;
        }
        public registrationcs()
        {
            InitializeComponent();
           
        }

        private void btn_update_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(uploadedFilePath) || package == null)
            {
                MessageBox.Show("Please upload an Excel file first.", "Error");
                return;
            }

            try
            {
                var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
                string enteredAadhaar = txt_AdharNo.Text; // Replace with the actual Aadhaar textbox name
                bool found = false;

                // Loop through the Excel rows to find the Aadhaar number
                for (int row = 2; row <= worksheet.Dimension.Rows; row++) // Start from row 2 (assuming row 1 is headers)
                {
                    string aadhaarInExcel = worksheet.Cells[row, 10].Text; // Column 15 for Aadhaar No (adjust if needed)

                    if (aadhaarInExcel == enteredAadhaar)
                    {
                        // Fetch corresponding fields from the Excel file
                        text_name.Text = worksheet.Cells[row, 1].Text.Length > 16 ? worksheet.Cells[row, 1].Text.Substring(0, 16) : worksheet.Cells[row, 1].Text; // Name (Column 1)
                        txt_fathername.Text = worksheet.Cells[row, 2].Text.Length > 32 ? worksheet.Cells[row, 2].Text.Substring(0, 32) : worksheet.Cells[row, 2].Text; // Father’s Name (Column 2)
                        txt_Dob.Text = worksheet.Cells[row, 3].Text.Length > 16 ? worksheet.Cells[row, 3].Text.Substring(0, 16) : worksheet.Cells[row, 3].Text; // D.O.B (Column 3)
                        txt_Doj.Text = worksheet.Cells[row, 4].Text.Length > 16 ? worksheet.Cells[row, 4].Text.Substring(0, 16) : worksheet.Cells[row, 4].Text; // D.O.J (Column 4)
                        txt_bloodgrp.Text = worksheet.Cells[row, 5].Text.Length > 16 ? worksheet.Cells[row, 5].Text.Substring(0, 16) : worksheet.Cells[row, 5].Text; // Blood Group (Column 5)
                        txt_mobileno.Text = worksheet.Cells[row, 6].Text.Length > 16 ? worksheet.Cells[row, 6].Text.Substring(0, 16) : worksheet.Cells[row, 6].Text ; // MOBILE NO (Column 7)
                        txt_empcode.Text = worksheet.Cells[row, 7].Text.Length > 16 ? worksheet.Cells[row,7].Text.Substring(0, 16) : worksheet.Cells[row,7 ].Text; // EMPLOYEE CODE (Column 8)
                        txt_email.Text = worksheet.Cells[row, 8].Text.Length > 32 ? worksheet.Cells[row, 8].Text.Substring(0, 32) : worksheet.Cells[row, 8].Text; // EMAIL (Column 9)
                        txt_empdprtment.Text = worksheet.Cells[row, 9].Text.Length > 16 ? worksheet.Cells[row, 9].Text.Substring(0, 16) : worksheet.Cells[row, 9].Text; ; // DEPARTMENT (Column 10)
                        txt_ehs.Text = worksheet.Cells[row, 11].Text.Length >16 ? worksheet.Cells[row, 11].Text.Substring(0, 16) : worksheet.Cells[row, 11].Text;// EHS NO (Column 11)16
                        txt_address.Text = worksheet.Cells[row, 12].Text.Length > 96 ? worksheet.Cells[row, 12].Text.Substring(0, 96) : worksheet.Cells[row, 12].Text;
    
                        txt_validupto.Text = worksheet.Cells[row, 13].Text.Length > 16 ? worksheet.Cells[row, 13].Text.Substring(0, 16) : worksheet.Cells[row, 13].Text; // VALID UPTO (Column 13)
                        txt_issueAuthority.Text = worksheet.Cells[row, 14].Text.Length > 32 ? worksheet.Cells[row, 14].Text.Substring(0, 32) : worksheet.Cells[row, 140].Text; // ISSUING AUTHORITY (Column 14)

                        found = true;
                        break;

                    }
                }

                if (!found)
                {
                    MessageBox.Show("Aadhaar number not found in the Excel file.", "Not Found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error");
            }
        }

        private void btn_upload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                uploadedFilePath = openFileDialog.FileName; // Store the file path in the class-level variable

                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    package = new ExcelPackage(new FileInfo(uploadedFilePath)); // Initialize ExcelPackage
                    MessageBox.Show("Excel file uploaded successfully.", "Success");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error");
                }
            }
        }

        private void Connect()
        {
            //=====================================================================
            // This function opens the port(connection) to ACR120 reader
            //=====================================================================
            // Variable declarations
            int ctr = 0;
            byte[] FirmwareVer = new byte[31];
            byte[] FirmwareVer1 = new byte[20];
            byte infolen = 0x00;
            string FirmStr;
            ACR120U.tReaderStatus ReaderStat = new ACR120U.tReaderStatus();

            if (g_isConnected)
            {
                // Device is already connected
                return;
            }

            g_rHandle = ACR120U.ACR120_Open(0);

            if (g_rHandle != 0)
            {
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_rHandle));
                tabstatusStrip.Text = "Reader Disconnected";
                button1.Enabled = false;
                return; // Exit function if connection fails
            }

            // Connection successful
            g_isConnected = true;
            tabstatusStrip.Text = "Reader Connected";
            button1.Enabled = true;

            // Get the DLL version the program is using
            g_retCode = ACR120U.ACR120_RequestDLLVersion(ref infolen, ref FirmwareVer[0]);
            if (g_retCode < 0)
            {
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
            }
            else
            {
                FirmStr = "";
                for (ctr = 0; ctr < Convert.ToInt16(infolen) - 1; ctr++)
                    FirmStr += char.ToString((char)(FirmwareVer[ctr]));
                // MessageBox.Show("DLL Version : " + FirmStr);
            }

            // Routine to get the firmware version
            g_retCode = ACR120U.ACR120_Status(g_rHandle, ref FirmwareVer1[0], ref ReaderStat);
            if (g_retCode < 0)
            {
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
            }
            else
            {
                FirmStr = "";
                for (ctr = 0; ctr < Convert.ToInt16(infolen); ctr++)
                    if ((FirmwareVer1[ctr] != 0x00) && (FirmwareVer1[ctr] != 0xFF))
                        FirmStr += char.ToString((char)(FirmwareVer1[ctr]));
                // MessageBox.Show("Firmware Version : " + FirmStr);
            }
        }


        private void btn_clear_Click(object sender, EventArgs e)
        {
            //errorProvider.Clear();

            ClearFields();
        }
        private void ClearFields()
        {
            // Clear all fields after verification
            text_name.Clear();
            txt_fathername.Clear();
            txt_Dob.Clear();
            txt_Doj.Clear();
            txt_bloodgrp.Clear();
            txt_mobileno.Clear();
            txt_empcode.Clear();
            txt_email.Clear();
            txt_empdprtment.Clear();
            txt_AdharNo.Clear();
            txt_ehs.Clear();
            txt_address.Clear();
            txt_validupto.Clear();
            txt_issueAuthority.Clear();
            lblProgress.Text = string.Empty;
            progressBar1.Value = 0;// Optionally, clear the progress label

            

        }


        private void button1_Click(object sender, EventArgs e)
        {
            WriteToSectors();
        }
        // Declare text1 as a TextBox control
        private void SelectCard()
        {
            //Variable Declarations
            byte[] ResultSN = new byte[11];
            byte ResultTag = 0x00;
            byte[] TagType = new byte[51];
            int ctr = 0;
            string SN = "";

            registrationcs.g_retCode = ACR120U.ACR120_Select(registrationcs.g_rHandle, ref TagType[0], ref ResultTag, ref ResultSN[0]);
            if (registrationcs.g_retCode < 0)

                MessageBox.Show("[X] " + ACR120U.GetErrMsg(registrationcs.g_retCode));

            else
            {
                if ((TagType[0] == 4) || (TagType[0] == 5))
                {

                    SN = "";
                    for (ctr = 0; ctr < 7; ctr++)
                    {
                        SN = SN + string.Format("{0:X2} ", ResultSN[ctr]);
                    }

                }
                else
                {

                    SN = "";
                    for (ctr = 0; ctr < ResultTag; ctr++)
                    {
                        SN = SN + string.Format("{0:X2} ", ResultSN[ctr]);
                    }

                }
                if (txt_AdharNo.Text == "")
                {
                    txt_AdharNo.Text = SN.Trim();
                }
            }
        }
      
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ReadData readForm = new ReadData(); // Create an instance of ReadForm
            readForm.ShowDialog(); // Show it as a popup (modal)
        }


        private void WriteToSectors()
        {
            bool isValid = true; // Flag to check if all mandatory fields are filled
            ErrorProvider errorProvider = new ErrorProvider(); // Create an error provider instance
            errorProvider.BlinkStyle = ErrorBlinkStyle.NeverBlink; // Stop blinking effect

            // Function to check and highlight empty fields
            void ValidateField(TextBox txtBox, string errorMessage)
            {
                if (string.IsNullOrWhiteSpace(txtBox.Text))
                {
                    isValid = false;
                    //txtBox.BackColor = Color.LightCoral; // Highlight textbox in red
                    errorProvider.SetIconAlignment(txtBox, ErrorIconAlignment.MiddleRight); // Show icon inside textbox
                    errorProvider.SetIconPadding(txtBox, -20); // Adjust icon position inside the textbox
                    errorProvider.SetError(txtBox, errorMessage); // Show error icon with message
                }
                else
                {
                    txtBox.BackColor = Color.White; // Reset color if field is filled
                    errorProvider.SetError(txtBox, ""); // Remove error
                }
            }

            // Validate mandatory fields
            ValidateField(txt_AdharNo, "Adhar No is required.");
            ValidateField(txt_empcode, "Employee ID is required.");
            ValidateField(text_name, "Name is required.");
            ValidateField(txt_mobileno, "Mobile No is required.");

            // Stop execution if any mandatory field is empty
            if (!isValid)
            {
                lblProgress.ForeColor = Color.Red;
                lblProgress.Text = "Please fill in all the mandatory fields.";
                return;
            }

            // Initialize progress tracking
            int totalWrites = 20; // This is based on the number of fields you are writing
            int currentWrite = 0;

            // Set progress bar maximum and initial value
            progressBar1.Value = 0;
            progressBar1.Maximum = totalWrites;
            progressBar1.Step = 1;
            lblProgress.Text = "0%";
            void UpdateProgress()
            {
                if (currentWrite < totalWrites)
                {
                    currentWrite++;
                    int percentage = (currentWrite * 100) / totalWrites;

                    // Ensure percentage does not exceed 100%
                    if (percentage > 100)
                        percentage = 100;

                    progressBar1.Value = currentWrite;
                    lblProgress.ForeColor = Color.Black;
                    lblProgress.Text = $"{percentage}%";
                    Application.DoEvents(); // Update UI immediately
                }
            }



            SelectCard(); 

            // Writing data and updating progress
            string data1 = text_name.Text.PadRight(32);
            Write(0, 1, data1.Substring(0, 16)); UpdateProgress();
            Write(0, 2, data1.Substring(16, 16)); UpdateProgress();

            string data2 = txt_fathername.Text.PadRight(32);
            Write(1, 0, data2.Substring(0, 16)); UpdateProgress();
            Write(1, 1, data2.Substring(16, 16)); UpdateProgress();

            string data3 = txt_Dob.Text.PadRight(16);
            Write(1, 2, data3); UpdateProgress();

            string data4 = txt_Doj.Text.PadRight(16);
            Write(2, 0, data4); UpdateProgress();

            string data5 = txt_bloodgrp.Text.PadRight(16);
            Write(2, 1, data5); UpdateProgress();

            string data6 = txt_mobileno.Text.PadRight(16);
            Write(2, 2, data6); UpdateProgress();

            string data7 = txt_empcode.Text.PadRight(16);
            Write(3, 1, data7); UpdateProgress();

            string data8 = txt_email.Text.PadRight(32);
            Write(4, 0, data8.Substring(0, 16)); UpdateProgress();
            Write(4, 1, data8.Substring(16, 16)); UpdateProgress();

            string data9 = txt_empdprtment.Text.PadRight(16);
            Write(4, 2, data9); UpdateProgress();

            string data10 = txt_AdharNo.Text.PadRight(16);
            Write(5, 0, data10); UpdateProgress();

            string data11 = txt_ehs.Text.PadRight(16);
            Write(5, 1, data11); UpdateProgress();

            string data12 = txt_address.Text.PadRight(96);
            Write(6, 0, data12.Substring(0, 16)); UpdateProgress();
            Write(6, 1, data12.Substring(16, 16)); UpdateProgress();
            Write(6, 2, data12.Substring(32, 16)); UpdateProgress();
            Write(7, 0, data12.Substring(48, 16)); UpdateProgress();
            Write(7, 1, data12.Substring(64, 16)); UpdateProgress();
            Write(7, 2, data12.Substring(80, 16)); UpdateProgress();
            


            string data13 = txt_validupto.Text.PadRight(16);
            Write(8, 0, data13); UpdateProgress();

            string data14 = txt_issueAuthority.Text.PadRight(32);
            Write(8, 1, data14.Substring(0, 16)); UpdateProgress();
            Write(8, 2, data14.Substring(16, 16)); UpdateProgress();

            // Ensure progress bar reaches 100%
            Application.DoEvents();
           // ReadCardData();
            bool result = VerifyWrittenData();
            if (result)
            {
                lblProgress.ForeColor = Color.DarkGreen;
                
                lblProgress.Text = "Data verification successful.";
                ClearFields();





            }
            else
            {
                progressBar1.Value = 0;
                lblProgress.ForeColor = Color.DarkRed;
                lblProgress.Text = "Data verification failed.";
            }


        }
        private void Write(int PhysicalSector, byte Block, string data)
        {
            
            long sto = 0;
            byte vKeyType = 0x00;
            int ctr, tmpInt = 0; 
           
            char[] charArray1 = new char[16];

            #region Write for Sector 0
            vKeyType = ACR120U.ACR120_LOGIN_KEYTYPE_A;
            //vKeyType = ACR120U.ACR120_LOGIN_KEYTYPE_STORED_A;
            //g_Sec = 0;
            
            PhysicalSector = Convert.ToInt16(PhysicalSector);
            tmpInt = Convert.ToInt16(Block);
            sto = 30;
            for (ctr = 0; ctr < 6; ctr++)
                g_pKey[ctr] = 0xFF;
            g_retCode = ACR120U.ACR120_Login(g_rHandle, Convert.ToByte(PhysicalSector), Convert.ToInt16(vKeyType),
                                         Convert.ToByte(sto), ref g_pKey[0]);
            if (g_retCode < 0)

                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));

            tmpInt = tmpInt + Convert.ToInt16(PhysicalSector) * 4;
            Block = Convert.ToByte(tmpInt);

            charArray1 = data.ToString().ToCharArray();
            byte[] dout = new byte[16];

            for (ctr = 0; ctr < 16; ctr++)
            {
                if (ctr < charArray1.Length)
                {
                    dout[ctr] = (byte)charArray1[ctr]; // Convert char to byte
                }
                else
                {
                    dout[ctr] = 0x00; // Fill the remaining bytes with 0 if the string is shorter than 16 characters
                }
            }


            g_retCode = ACR120U.ACR120_Write(g_rHandle, Block, ref dout[0]);

            if (g_retCode < 0)
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
            #endregion
        }
        private bool ReadCardData()
        {
            bool chk = false;

            // Read data from the card and populate the text fields
            string name = Read(0, 1) + Read(0, 2).Trim();
            if (text_name.Text != name)
            {
                return chk;
            }

            string fatherName = Read(1, 0) + Read(1, 1).Trim();
            if (txt_fathername.Text != fatherName)
            {
                return chk;
            }

            string dob = Read(1, 2).Trim();
            if (txt_Dob.Text != dob)
            {
                return chk;
            }

            string doj = Read(2, 0).Trim();
            if (txt_Doj.Text != doj)
            {
                return chk;
            }

            string bloodGroup = Read(2, 1).Trim();
            if (txt_bloodgrp.Text != bloodGroup)
            {
                return chk;
            }

            string mobileNo = Read(2, 2).Trim();
            if (txt_mobileno.Text != mobileNo)
            {
                return chk;
            }

            string empCode = Read(3, 1).Trim();
            if (txt_empcode.Text != empCode)
            {
                return chk;
            }

            string email = Read(4, 0) + Read(4, 1).Trim();
            if (txt_email.Text != email)
            {
                return chk;
            }

            string empDepartment = Read(4, 2).Trim();
            if (txt_empdprtment.Text != empDepartment)
            {
                return chk;
            }

            string adharNo = Read(5, 0).Trim();
            if (txt_AdharNo.Text != adharNo)
            {
                return chk;
            }

            string ehs = Read(5, 1).Trim();
            if (txt_ehs.Text != ehs)
            {
                return chk;
            }

            string address = (Read(6, 0) + Read(6, 1) + Read(6, 2) + Read(7,0)+ Read(7, 1)+ Read(7, 2)).Trim();
            if (txt_address.Text != address)
            {
                return chk;
            }

            string validUpto = Read(8, 0).Trim();
            if (txt_validupto.Text != validUpto)
            {
                return chk;
            }

            string issueAuthority = Read(8, 1) + Read(8, 2).Trim();
            if (txt_issueAuthority.Text != issueAuthority)
            {
                return chk;
            }

            chk = true;
            return chk;
        }


        private bool VerifyWrittenData()
        {
            // Read the data from the card after writing
            ReadCardData();

            // Store read values in variables
            string name = (Read(0, 1) + Read(0, 2)).Trim();
            string fatherName = (Read(1, 0) + Read(1, 1)).Trim();
            string dob = Read(1, 2).Trim();
            string doj = Read(2, 0).Trim();
            string bloodGroup = Read(2, 1).Trim();
            string mobileNo = Read(2, 2).Trim();
            string empCode = Read(3, 1).Trim();
            string email = (Read(4, 0) + Read(4, 1)).Trim();
            string empDepartment = Read(4, 2).Trim();
            string adharNo = Read(5, 0).Trim();
            string ehs = Read(5, 1).Trim();
            string address = (Read(6, 0) + Read(6, 1) + Read(6, 2)+ Read(7,0)+ Read(7, 1)+ Read(7, 2)).Trim();
            string validUpto = Read(8, 0).Trim();
            string issueAuthority = (Read(8, 1) + Read(8, 2)).Trim();

            // Compare each field of the written data with the read data
            bool isMatch = IsDataMatching(text_name.Text, name) &&
                           IsDataMatching(txt_fathername.Text, fatherName) &&
                           IsDataMatching(txt_Dob.Text, dob) &&
                           IsDataMatching(txt_Doj.Text, doj) &&
                           IsDataMatching(txt_bloodgrp.Text, bloodGroup) &&
                           IsDataMatching(txt_mobileno.Text, mobileNo) &&
                           IsDataMatching(txt_empcode.Text, empCode) &&
                           IsDataMatching(txt_email.Text, email) &&
                           IsDataMatching(txt_empdprtment.Text, empDepartment) &&
                           IsDataMatching(txt_AdharNo.Text, adharNo) &&
                           IsDataMatching(txt_ehs.Text, ehs) &&
                           IsDataMatching(txt_address.Text, address) &&
                           IsDataMatching(txt_validupto.Text, validUpto) &&
                           IsDataMatching(txt_issueAuthority.Text, issueAuthority);

            return isMatch;
           

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private bool IsDataMatching(string writtenData, string readData)
        {
            // Trim and compare written and read data without considering case differences
            return string.Equals(writtenData?.Trim(), readData?.Trim(), StringComparison.OrdinalIgnoreCase);
        }





        private string Read(int PhysicalSector, byte Blck2)
        {
            byte[] PassRead = new byte[16];
            long sto = 0;
            byte vKeyType = 0x00;
            int ctr, tmpInt2 = 2;
            string CardBalance = "";

            vKeyType = ACR120U.ACR120_LOGIN_KEYTYPE_A;
            //vKeyType1 = ACR120U.ACR120_LOGIN_KEYTYPE_STORED_A;
            PhysicalSector = Convert.ToInt16(PhysicalSector);
            tmpInt2 = Convert.ToInt16(Blck2);
            sto = 30;
            for (ctr = 0; ctr < 6; ctr++)
                g_pKey[ctr] = 0xFF;
            g_retCode = ACR120U.ACR120_Login(g_rHandle, Convert.ToByte(PhysicalSector), Convert.ToInt16(vKeyType),
                                             Convert.ToByte(sto), ref g_pKey[0]);
            if (g_retCode < 0)
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
            tmpInt2 = tmpInt2 + Convert.ToInt16(PhysicalSector) * 4;

            Blck2 = Convert.ToByte(tmpInt2);

            g_retCode = ACR120U.ACR120_Read(g_rHandle, Blck2, ref PassRead[0]);
            if (g_retCode < 0)
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
            else
            {
                dstrpass1 = "";
                for (ctr = 0; ctr < 16; ctr++)
                {
                    dstrpass1 = dstrpass1 + char.ToString((char)(PassRead[ctr]));
                }
                CardBalance = Convert.ToString(dstrpass1);
                if (CardBalance == "\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0")
                    CardBalance = "0";
            }
            return CardBalance;
        }
    }
}