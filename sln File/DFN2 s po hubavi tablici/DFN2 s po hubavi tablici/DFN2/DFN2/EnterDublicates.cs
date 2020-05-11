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
    public partial class EnterDublicatesForm : Form
    {
        private List<List<int>> listOfLists { get; set; }
        private List<int> startingList;
        public EnterDublicatesForm()
        {
            InitializeComponent();
        }
        public EnterDublicatesForm(string nomNumber)
        {
            InitializeComponent();
            NomenclatureNumberLabel.Text = nomNumber;
            startingList = new List<int>();
            listOfLists = new List<List<int>>();
        }

        private void AddAvailableDupSetButton_Click(object sender, EventArgs e)
        {
            if (Int32.TryParse(StartAvailableDupTextBox.Text, out int result) == false || Int32.TryParse(EndAvailableDupTextBox.Text, out result) == false || int.Parse(StartAvailableDupTextBox.Text) > int.Parse(EndAvailableDupTextBox.Text))
            {
                warningLabel.Text="Incorrect data";
                warningLabel.Visible=true;
            }
            else
            {
                int numberOfElements = 0;

                Helper.AddRange(int.Parse(StartAvailableDupTextBox.Text), int.Parse(EndAvailableDupTextBox.Text), startingList);
                Helper.CutList(startingList, listOfLists);
                AvailableDupSetListBox.Items.Clear();
                numberOfElements = int.Parse(AvailableDupSetCouterLabel.Text) + int.Parse(EndAvailableDupTextBox.Text) - int.Parse(StartAvailableDupTextBox.Text) + 1;
                StartAvailableDupTextBox.Text = "";
                EndAvailableDupTextBox.Text = "";
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                }
                AvailableDupSetCouterLabel.Text = numberOfElements.ToString();
            }
        }

        private void ConfirmDataButton_Click(object sender, EventArgs e)
        {
            int[] fromToArray = new int[2];
            int from = 0;
            int to = 0;
            foreach (string item in AvailableDupSetListBox.Items)
            {
                fromToArray = item.Split('-').Select(int.Parse).ToArray();
                from = fromToArray[0];
                to = fromToArray[1];
                int numberOfIterations = to - from + 1;
                for (int i = 0; i < numberOfIterations; i++)
                {
                    Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                    dataAccess.InsertDocument(Helper.MakeTo6((from + i).ToString()), NomenclatureNumberLabel.Text, "2018", "Наличен");
                }
            }
            foreach (string item in DupRegBookListBox.Items)
            {
                fromToArray = item.Split('-').Select(int.Parse).ToArray();
                from = fromToArray[0];
                to = fromToArray[1];
                int numberOfIterations = to - from + 1;
                for (int i = 0; i < numberOfIterations; i++)
                {
                    Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                    dataAccess.InsertDocument(Helper.MakeTo6((from + i).ToString()), NomenclatureNumberLabel.Text, "2018", "Издаден по регистрационна книга");
                }
            }
            foreach (string item in ReadyForDestructionListBox.Items)
            {
                fromToArray = item.Split('-').Select(int.Parse).ToArray();
                from = fromToArray[0];
                to = fromToArray[1];
                int numberOfIterations = to - from + 1;
                for (int i = 0; i < numberOfIterations; i++)
                {
                    Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                    dataAccess.InsertDocument(Helper.MakeTo6((from + i).ToString()), NomenclatureNumberLabel.Text, "2018", "Годен за унищожаване");
                }
            }
            foreach (string item in CanceledDupListBox.Items)
            {
                fromToArray = item.Split('-').Select(int.Parse).ToArray();
                from = fromToArray[0];
                to = fromToArray[1];
                int numberOfIterations = to - from + 1;
                for (int i = 0; i < numberOfIterations; i++)
                {
                    Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                    dataAccess.InsertDocument(Helper.MakeTo6((from + i).ToString()), NomenclatureNumberLabel.Text, "2018", "Анулиран");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Int32.TryParse(StartDupRegBookTextBox.Text, out int result) == false || Int32.TryParse(EndDupRegBookTextBox.Text, out result) == false || int.Parse(StartDupRegBookTextBox.Text) > int.Parse(EndDupRegBookTextBox.Text))
            {
                warningLabel.Text="Incorrect data";
                warningLabel.Visible=true;
            }
            else
            {
                int from = int.Parse(StartDupRegBookTextBox.Text);
                int to = int.Parse(EndDupRegBookTextBox.Text);
                StartDupRegBookTextBox.Text = "";
                EndDupRegBookTextBox.Text = "";
                int numberOfElements = to - from + 1;
                listOfLists.Clear();
                Helper.RemoveRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                AvailableDupSetListBox.Items.Clear();
                DupRegBookListBox.Items.Add(Helper.MakeTo6(from.ToString()) + "-" + Helper.MakeTo6(to.ToString()));
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));

                }
                DupRegBookCounterLabel.Text = (numberOfElements + int.Parse(DupRegBookCounterLabel.Text)).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) - numberOfElements).ToString();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (DupRegBookListBox.SelectedItem == null)
            {
                warningLabel.Visible=true;
                warningLabel.Text="Please select a range first";
            }
            else
            {
                string range = DupRegBookListBox.SelectedItem.ToString();
                DupRegBookListBox.Items.Remove(DupRegBookListBox.SelectedItem);
                int[] fromToArray = range.Split('-').Select(int.Parse).ToArray();
                int from = fromToArray[0];
                int to = fromToArray[1];
                int numberOfElements = to - from + 1;
                AvailableDupSetListBox.Items.Clear();
                listOfLists.Clear();
                Helper.ResetRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                }
                DupRegBookCounterLabel.Text = (int.Parse(DupRegBookCounterLabel.Text) - numberOfElements).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) + numberOfElements).ToString();
            }
        }

        private void AddAvailableDupButton_Click(object sender, EventArgs e)
        {
            if (Int32.TryParse(StartReadyForDesTextBox.Text, out int result) == false || Int32.TryParse(EndReadyForDestructionTextBox.Text, out result) == false || int.Parse(StartReadyForDesTextBox.Text) > int.Parse(EndReadyForDestructionTextBox.Text))
            {
                warningLabel.Text="Incorrect data";
                warningLabel.Visible=true;
            }
            else
            {
                int from = int.Parse(StartReadyForDesTextBox.Text);
                int to = int.Parse(EndReadyForDestructionTextBox.Text);
                StartReadyForDesTextBox.Text = "";
                EndReadyForDestructionTextBox.Text = "";
                int numberOfElements = to - from + 1;
                listOfLists.Clear();
                Helper.RemoveRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                AvailableDupSetListBox.Items.Clear();
                ReadyForDestructionListBox.Items.Add(Helper.MakeTo6(from.ToString()) + "-" + Helper.MakeTo6(to.ToString()));
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                }
                AvailableDupCounterLabel.Text = (int.Parse(AvailableDupCounterLabel.Text) + numberOfElements).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) - numberOfElements).ToString();
            }
        }

        private void RemoveAvailableDupButton_Click(object sender, EventArgs e)
        {
            if (ReadyForDestructionListBox.SelectedItem == null)
            {
                warningLabel.Visible=true;
                warningLabel.Text="Please select a range first";
            }
            else
            {
                string range = ReadyForDestructionListBox.SelectedItem.ToString();
                ReadyForDestructionListBox.Items.Remove(ReadyForDestructionListBox.SelectedItem);
                int[] fromToArray = range.Split('-').Select(int.Parse).ToArray();
                int from = fromToArray[0];
                int to = fromToArray[1];
                int numberOfElements = to - from + 1;
                AvailableDupSetListBox.Items.Clear();
                listOfLists.Clear();
                Helper.ResetRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));

                }
                AvailableDupCounterLabel.Text = (int.Parse(AvailableDupCounterLabel.Text) - numberOfElements).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) + numberOfElements).ToString();
            }
        }

        private void AddCanceledDupButton_Click(object sender, EventArgs e)
        {
            if (Int32.TryParse(StartCanceledDupTextBox.Text, out int result) == false || Int32.TryParse(EndCanceledDupTextBox.Text, out result) == false || int.Parse(StartCanceledDupTextBox.Text) > int.Parse(EndCanceledDupTextBox.Text))
            {
                warningLabel.Text="Incorrect data";
                warningLabel.Visible=true;
            }
            else
            {
                int from = int.Parse(StartCanceledDupTextBox.Text);
                int to = int.Parse(EndCanceledDupTextBox.Text);
                StartCanceledDupTextBox.Text = "";
                EndCanceledDupTextBox.Text = "";
                int numberOfElements = to - from + 1;
                listOfLists.Clear();
                Helper.RemoveRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                AvailableDupSetListBox.Items.Clear();
                CanceledDupListBox.Items.Add(Helper.MakeTo6(from.ToString()) + "-" + Helper.MakeTo6(to.ToString()));
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                }
                CanceledDupCounterLabel.Text = (int.Parse(CanceledDupCounterLabel.Text) + numberOfElements).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) - numberOfElements).ToString();
            }
        }

        private void RemoveCanceledDupButton_Click(object sender, EventArgs e)
        {
            if (CanceledDupListBox.SelectedItem == null)
            {
                warningLabel.Visible=true;
                warningLabel.Text="Please select a range first";
            }
            else
            {
                string range = CanceledDupListBox.SelectedItem.ToString();
                CanceledDupListBox.Items.Remove(CanceledDupListBox.SelectedItem);
                int[] fromToArray = range.Split('-').Select(int.Parse).ToArray();
                int from = fromToArray[0];
                int to = fromToArray[1];
                int numberOfElements = to - from + 1;
                AvailableDupSetListBox.Items.Clear();
                listOfLists.Clear();
                Helper.ResetRange(from, to, startingList);
                Helper.CutList(startingList, listOfLists);
                foreach (var list in listOfLists)
                {
                    AvailableDupSetListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));

                }
                CanceledDupCounterLabel.Text = (int.Parse(CanceledDupCounterLabel.Text) - numberOfElements).ToString();
                AvailableDupSetCouterLabel.Text = (int.Parse(AvailableDupSetCouterLabel.Text) + numberOfElements).ToString();
            }
        }

        private void AddForDestructionButton_Click(object sender, EventArgs e)
        {
            AvailableDupSetListBox.Items.Clear();
            int numberOfElements = 0;
            foreach (var list in listOfLists)
            {
                ReadyForDestructionListBox.Items.Add(Helper.MakeTo6(list[0].ToString()) + "-" + Helper.MakeTo6(list[list.Count - 1].ToString()));
                numberOfElements += list.Count;
            }
            AvailableDupCounterLabel.Text = (numberOfElements + int.Parse(AvailableDupCounterLabel.Text)).ToString();
            AvailableDupSetCouterLabel.Text = "0";
        }
    }
}
