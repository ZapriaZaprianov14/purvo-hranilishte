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
    public partial class ExtraGivenDuplicationForm : Form
    {
        private TablesForm tables;
        private List<int> startingList;
        private List<List<int>> listOfLists;
        public ExtraGivenDuplicationForm()
        {
            InitializeComponent();
        }
        public ExtraGivenDuplicationForm(string nomNumber)
        {
            InitializeComponent();
            NomenclatureNumberLabel.Text = nomNumber;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Int32.TryParse(StartNumberTextBox.Text, out int result) == false || Int32.TryParse(EndNumberTextBox.Text, out result) == false || int.Parse(StartNumberTextBox.Text) > int.Parse(EndNumberTextBox.Text))
            {
                //warningLabel.Text="Incorrect data";
                //warning.Visible=true;
            }
            else
            {
                int numberOfElements = 0;

                Helper.AddRange(int.Parse(StartNumberTextBox.Text), int.Parse(EndNumberTextBox.Text), startingList);
                Helper.CutList(startingList, listOfLists);
                DocumentsLitBox.Items.Clear();
                numberOfElements = int.Parse(docsCounter.Text) + int.Parse(EndNumberTextBox.Text) - int.Parse(StartNumberTextBox.Text) + 1;
                StartNumberTextBox.Text = "";
                EndNumberTextBox.Text = "";
                foreach (var list in listOfLists)
                {
                    DocumentsLitBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                }
                docsCounter.Text = numberOfElements.ToString();
                listOfLists.Clear();
            }
        }

        private void RemoveAvailableDupButton_Click(object sender, EventArgs e)
        {

        }
    }
}
