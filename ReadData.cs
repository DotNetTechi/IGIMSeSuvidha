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
    public partial class ReadData : Form
    {
        public static int g_rHandle, g_retCode;
        public static bool g_isConnected = false;
        public static byte g_Sec;
        public static byte[] g_pKey = new byte[6];
        ComboBox cbPort = new ComboBox();
        int iTips;
        int iAttempt;
        string dstrpass1;
        public ReadData()
        {
            InitializeComponent();
        }
        private void Connect()
        {
            int ctr = 0;
            byte[] FirmwareVer = new byte[31];
            byte[] FirmwareVer1 = new byte[20];
            byte infolen = 0x00;
            string FirmStr;
            ACR120U.tReaderStatus ReaderStat = new ACR120U.tReaderStatus();

            if (g_isConnected)
            {
                //MessageBox.Show("Device is already connected.");
                return;
            }

            g_rHandle = ACR120U.ACR120_Open(0);
            if (g_rHandle != 0)
            {
                MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_rHandle));
                btn_read.Enabled = false;  // Disable Read button if connection fails
            }
            else
            {
                g_isConnected = true;
                btn_read.Enabled = true;  // Enable Read button when connected

                g_retCode = ACR120U.ACR120_RequestDLLVersion(ref infolen, ref FirmwareVer[0]);
                if (g_retCode < 0)
                    MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
                else
                {
                    FirmStr = "";
                    for (ctr = 0; ctr < Convert.ToInt16(infolen) - 1; ctr++)
                        FirmStr += char.ToString((char)(FirmwareVer[ctr]));
                }

                g_retCode = ACR120U.ACR120_Status(g_rHandle, ref FirmwareVer1[0], ref ReaderStat);
                if (g_retCode < 0)
                    MessageBox.Show("[X] " + ACR120U.GetErrMsg(g_retCode));
                else
                {
                    FirmStr = "";
                    for (ctr = 0; ctr < Convert.ToInt16(infolen); ctr++)
                        if ((FirmwareVer1[ctr] != 0x00) && (FirmwareVer1[ctr] != 0xFF))
                            FirmStr += char.ToString((char)(FirmwareVer1[ctr]));
                }
            }
        }


        private void btn_clear_Click(object sender, EventArgs e)
        {

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
            txt_validupto.Clear();
            txt_address.Clear();
            txt_issueAuthority.Clear();
            txt_validupto.Clear();
            CSVStatus_lbl.Text = string.Empty;
            txtcard_id.Clear();


           
        }

        private void ReadData_Load(object sender, EventArgs e)
        {
            Connect();
        }
        private void SelectCard()
        {
            //Variable Declarations
            byte[] ResultSN = new byte[11];
            byte ResultTag = 0x00;
            byte[] TagType = new byte[51];
            int ctr = 0;
            string SN = "";

            ReadData.g_retCode = ACR120U.ACR120_Select(ReadData.g_rHandle, ref TagType[0], ref ResultTag, ref ResultSN[0]);
            if (ReadData.g_retCode < 0)

                MessageBox.Show("[X] " + ACR120U.GetErrMsg(ReadData.g_retCode));

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
                if (txtcard_id.Text == "")
                {
                    txtcard_id.Text = SN.Trim();
                }
            }
        }

        private void ReadCardData()
        {
            SelectCard();

            string name = Read(0, 1) + Read(0, 2).Trim();
            text_name.Text = !string.IsNullOrEmpty(name) ? name : "No Data";

            string fatherName = Read(1, 0) + Read(1, 1).Trim();
            txt_fathername.Text = !string.IsNullOrEmpty(fatherName) ? fatherName: "No Data";

            txt_Dob.Text = Read(1, 2).Trim();
            txt_Doj.Text = Read(2, 0).Trim();
            txt_bloodgrp.Text = Read(2, 1).Trim();
            txt_mobileno.Text = Read(2, 2).Trim();
            txt_empcode.Text = Read(3, 1).Trim();

            string email = Read(4, 0) + Read(4, 1).Trim();
            txt_email.Text = !string.IsNullOrEmpty(email) ? email: "No Data";

            txt_empdprtment.Text = Read(4, 2).Trim();
            txt_AdharNo.Text = Read(5, 0).Trim();
            txt_ehs.Text = Read(5, 1).Trim();

            string address = (Read(6, 0) + Read(6, 1) + Read(6, 2) + Read(7, 0) + Read(7, 1) + Read(7, 2)+Read(8,0)).Trim();
            txt_address.Text = !string.IsNullOrEmpty(address) ? address : "No Data";

            txt_validupto.Text = Read(9, 0).Trim();

            string issueAuthority = Read(9, 1) + Read(9, 2).Trim();
            txt_issueAuthority.Text = !string.IsNullOrEmpty(issueAuthority) ? issueAuthority : "No Data";

            // Now create the CSV file and update the status strip
            bool isCsvCreated = CreateCSV(name, fatherName, txt_Dob.Text, txt_Doj.Text, txt_bloodgrp.Text, txt_mobileno.Text,
                                          txt_empcode.Text, email, txt_empdprtment.Text, txt_AdharNo.Text, txt_ehs.Text,
                                          address, txt_validupto.Text, issueAuthority);

            if (isCsvCreated)
            {
                CSVStatus_lbl.Text = "CSV file created successfully.";
            }
            else
            {
                CSVStatus_lbl.Text = "Failed to create CSV file.";
            }
        }

        private bool CreateCSV(string name, string fatherName, string dob, string doj, string bloodGroup,
                             string mobileNo, string empCode, string email, string department,
                             string aadharNo, string ehs, string address, string validUpto,
                             string issueAuthority)
        {
            string filePath = "ReadLog.csv";
            bool fileExists = File.Exists(filePath);
            StreamWriter writer = null;

            try
            {
                writer = new StreamWriter(filePath, true, System.Text.Encoding.UTF8);

                if (!fileExists)
                {
                    writer.WriteLine("Name,Father Name,DOB,DOJ,Blood Group,Mobile No,Emp Code,Email,Department,Aadhar No,EHS,Address,Valid Upto,Issue Authority");
                }

                // Escape double quotes by replacing them with two double quotes
                string FormatField(string field) => $"\"{field.Replace("\"", "\"\"")}\"";

                writer.WriteLine($"{FormatField(name)},{FormatField(fatherName)},{FormatField(dob)},{FormatField(doj)},{FormatField(bloodGroup)},{FormatField(mobileNo)},{FormatField(empCode)},{FormatField(email)},{FormatField(department)},{FormatField(aadharNo)},{FormatField(ehs)},{FormatField(address)},{FormatField(validUpto)},{FormatField(issueAuthority)}");

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                writer?.Close();
            }
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
        
        private void btn_read_Click(object sender, EventArgs e)
        {
            ReadCardData();
        }

     

    }
    }


