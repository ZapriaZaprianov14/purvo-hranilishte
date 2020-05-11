using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DFN2
{
    public partial class FirstStepsDialog : Form
    {
        public FirstStepsDialog()
        {
            InitializeComponent();
        }

        private void HomeDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SchoolInfo_Click(object sender, EventArgs e)
        {
            label1.Visible = true;
            button5.Visible = true;
            button6.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SchoolInfoForm schoolInfoForm = new SchoolInfoForm();
            this.Hide();
            schoolInfoForm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            CommissionInfoForm commissionInfoForm = new CommissionInfoForm();
            this.Hide();
            commissionInfoForm.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(OrderNumberTB.Text=="")
            {
                warningLabel.Visible = true;
                warningLabel.Text = "Invalid data";

            }
            else
            {
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.InsertOrder(OrderNumberTB.Text, OrderDataPicker.Value);
                warningLabel.Visible = true;
                warningLabel.Text = "Data successfuly inserted";
                Globals.OrderNumber = OrderNumberTB.Text;
                Globals.OrderDate = OrderDataPicker.Value;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TablesForm tablesForm = new TablesForm();
            this.Hide();
            tablesForm.ShowDialog();
        }
    }
}
