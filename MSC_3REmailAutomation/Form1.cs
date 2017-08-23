using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace MSC_3REmailAutomation
{
    

    public partial class Form1 : Form
    {
        //private delegate void SetDGVValueDelegate(BindingList<Something> items);

        public static DataTable dt;
        public static double  v = 0;
        public Form1()
        {
            InitializeComponent();
           
            dt = new DataTable();

            myWorker.DoWork += new DoWorkEventHandler(myWorker_DoWork);
            myWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(myWorker_RunWorkerCompleted);
            myWorker.ProgressChanged += new ProgressChangedEventHandler(myWorker_ProgressChanged);
            myWorker.WorkerReportsProgress = true;
            myWorker.WorkerSupportsCancellation = true;

            backWorker.DoWork += new DoWorkEventHandler(backWorker_DoWork);
            backWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backWorker_RunWorkerCompleted);
            backWorker.ProgressChanged += new ProgressChangedEventHandler(backWorker_ProgressChanged);
            backWorker.WorkerReportsProgress = true;
            backWorker.WorkerSupportsCancellation = true;


           


            label3.Visible = false;
            label4.Visible = false;


        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                timer1.Interval = 1000;
                timer1.Enabled = true;
                timer1.Tick += new System.EventHandler(timer1_Tick);
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();

                if (!myWorker.IsBusy)
                {
                    myWorker.RunWorkerAsync();
                }
                else
                {
                    MessageBox.Show("Data retrieval in progress");
                }
                label3.Visible = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!backWorker.IsBusy)
                {
                    backWorker.RunWorkerAsync();
                }
                else
                {
                    MessageBox.Show("Email Send in progress");
                }
                label4.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
        private void SetDGVValue()
        {
            try
            {
                dt = SendEmail.GetDataSet(Properties.Settings.Default.connString, SendEmail.getQuery()).Tables[0]; 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           

        }
        private void SetDGVValue1()
        {
            //dataGridView1.Columns.Clear();
            try
            {
                dataGridView1.DataSource = dt;
                SendEmail.AddOutOfOfficeColumn(dataGridView1);
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells["checkBoxColumn"].Value = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        
        
        protected void myWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                dataGridView1.Columns.Clear();
                SetDGVValue();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

        }

       
        protected void myWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (!e.Cancelled && e.Error == null)//Check if the worker has been canceled or if an error occurred
                {
                    string result = (string)e.Result;//Get the result from the background thread
                    //label3.Text = result;
                    SetDGVValue1();
                }
                else if (e.Cancelled)
                {
                    label3.Text = "User Canceled";
                }
                else
                {
                    label3.Text = "An error has occurred";
                }
                label3.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
            //label3.Visible = false;
            //myWorker.Dispose();
        }

        protected void myWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                label3.Text = " Progress: - " + e.ProgressPercentage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //v = v + 0.5 ; 
            //label4.Text = "Fetching data ...Time Elapsed-" + (v).ToString();
        }



        protected void backWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string message = string.Empty;
                
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    bool isSelected = Convert.ToBoolean(row.Cells["checkBoxColumn"].Value);
                    if (isSelected)
                    {
                        message += Environment.NewLine;
                        message += row.Cells["MST_ID"].Value.ToString();
                        int i = Int32.Parse(row.Cells["Gap of Days"].Value.ToString());
                        if ((i == 2 || i == 5 | i == 8))
                        {


                            SendEmail.sendExchangeEmail(SendEmail.getEmail(i, row.Cells["MarketingRequestName"].Value.ToString(),
                            row.Cells["MR_ID"].Value.ToString(),
                            row.Cells["MST_ID"].Value.ToString(),
                            row.Cells["MS_ID"].Value.ToString(),
                            row.Cells["MR_Title"].Value.ToString(),
                            row.Cells["Title"].Value.ToString(),
                            row.Cells["EmailAddress"].Value.ToString(),
                            row.Cells["SO_SubmittedDate_SubTimezone"].Value.ToString(),
                            row.Cells["DueDate_SubTimezone"].Value.ToString(),
                            row.Cells["SubsidiaryName"].Value.ToString(),
                            row.Cells["areaname"].Value.ToString(),
                            row.Cells["Program"].Value.ToString(),
                            row.Cells["ServiceTypeName"].Value.ToString(),
                            row.Cells["EPCampaignName"].Value.ToString(),
                            row.Cells["Link"].Value.ToString()
                            ), 
                            row.Cells["MR_ID"].Value.ToString(),
                            row.Cells["MR_Title"].Value.ToString(),
                            row.Cells["MarketingRequestOwner"].Value.ToString() + ";" + row.Cells["IRRN_Aliases"].Value.ToString(),
                            row.Cells["BuilderEmail"].Value.ToString() + ";" + row.Cells["SOP_Email"].Value.ToString() + ";" + row.Cells["FactoryLeadEmail"].Value.ToString(),
                            i
                            );
                            //[MarketingRequestName],[MR_ID],
                        }
                    }
                }

               // MessageBox.Show("Selected Values" + message);
                //SendEmail.sendExchangeEmail(SendEmail.getCSS() + SendEmail.getHeader() + SendEmail.getName("Paul Pogba") + SendEmail.getFooter());
                // SendEmail.sendExchangeEmail(SendEmail.getEmail(1));
                MessageBox.Show("Emails sent successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        protected void backWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

           // label4.Text = " Progress: - " + e.ProgressPercentage;
        }
        protected void backWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (!e.Cancelled && e.Error == null)//Check if the worker has been canceled or if an error occurred
                {
                    string result = (string)e.Result;//Get the result from the background thread
                    //label3.Text = result;
                    //SetDGVValue1();
                }
                else if (e.Cancelled)
                {
                    label4.Text = "User Canceled";
                }
                else
                {
                    label4.Text = "An error has occurred";
                }
                label4.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
            //myWorker.Dispose();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
      
    }
}
