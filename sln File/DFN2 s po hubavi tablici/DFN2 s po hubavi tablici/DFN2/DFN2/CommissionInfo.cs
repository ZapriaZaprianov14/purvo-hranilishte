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
    public partial class CommissionInfoForm : Form
    {
        public CommissionInfoForm()
        {
            InitializeComponent();
            Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
            bool managerInserted = dataAccess.CheckIfManagerInserted();
            if (managerInserted)
            {
                UpdateButton.Visible = true;
                UpdateButton.Enabled = true;
                ConfirmButton.Visible = false;
                ConfirmButton.Enabled = false;
            }
            else
            {
                UpdateButton.Visible = false;
                ConfirmButton.Visible = true;
                UpdateButton.Enabled = false;
            }
        }

        private void CommissionInfoForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        #region Menu Buttons
        private void HomeDialogButton_Click(object sender, EventArgs e)
        {
            FirstStepsDialog homeDialog = new FirstStepsDialog();
            this.Hide();
            homeDialog.ShowDialog();
        }

        private void SchoolInfoButton_Click(object sender, EventArgs e)
        {
            SchoolInfoForm schoolInfoForm = new SchoolInfoForm();
            this.Hide();
            schoolInfoForm.ShowDialog();
        }

        private void DuplicatesButton_Click(object sender, EventArgs e)
        {
            DuplicateTimer.Start();
        }

        #region Duplicate Number Button Click
        private void Dup1Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-30a");
            enterDublicatesForm.Show();
        }

        private void Dup2Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-22a");
            enterDublicatesForm.Show();
        }

        private void Dup3Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-44a");
            enterDublicatesForm.Show();
        }

        private void Dup4Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-54a");
            enterDublicatesForm.Show();
        }

        private void Dup5Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-54aB");
            enterDublicatesForm.Show();
        }

        private void Dup6Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm("3-27aB");
            enterDublicatesForm.Show();
        }
        #endregion

        private void TablesButton_Click(object sender, EventArgs e)
        {
            TablesForm tablesForm = new TablesForm();
            this.Hide();
            tablesForm.ShowDialog();
        }

        private void ReferencesButton_Click(object sender, EventArgs e)
        {
            ReferencesTimer.Start();
        }

        #region References Buttons 
        private void RestDupButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Остатък дубликати от минали години");
            referencesForm.Show();
        }

        private void GivenButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Предадени на други институции");
            referencesForm.Show();
        }

        private void ReceivedByOtherButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Приети от други институции");
            referencesForm.Show();
        }

        private void ReceivedByRequestButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Приети по заявка");
            referencesForm.Show();
        }

        private void RegBookButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Издадени по регистрационна книга");
            referencesForm.Show();
        }
        #endregion
        private void MakingDocumentsButton_Click(object sender, EventArgs e)
        {
            MakingDocumentsForm makingDocuments = new MakingDocumentsForm();
            this.Hide();
            makingDocuments.ShowDialog();
        }
        #endregion

        #region Duplicate Timer Tick Function - using it like counter for duplicate drop menu

        // using this variable to check if the menu is droped or not
        private bool notDroped = true;
        private void DuplicateTimer_Tick(object sender, EventArgs e)
        {
            // if the panel is not droped to become larger with 10 pixels 
            if (notDroped)
            {
                DuplicateDropPanel.Height += 10;

                // if the size of the panes is eaqual to its maximum size to stop the timer and report it as "Droped"
                if (DuplicateDropPanel.Size == DuplicateDropPanel.MaximumSize)
                {
                    DuplicateTimer.Stop();
                    notDroped = false;
                }
            }

            // if the panel is droped to become smaller with 10 pixels
            else
            {
                DuplicateDropPanel.Height -= 10;

                // if the size of the panes is eaqual to its minimum size to stop the timer and report it as "Not droped"
                if (DuplicateDropPanel.Size == DuplicateDropPanel.MinimumSize)
                {
                    DuplicateTimer.Stop();
                    notDroped = true;
                }
            }
        }

        #endregion

        #region References Timer Function - using it like counter for reference drop menu

        // using this variable to check if the menu is droped or not
        private bool droped = false;
        private void ReferencesTimer_Tick(object sender, EventArgs e)
        {
            // if the panel is NOT droped to become larger with 10 pixels 
            if (!droped)
            {
                ReferenceDropPanel.Height += 10;

                // if the size of the panes is eaqual to its maximum size to stop the timer and report it as "Droped"
                if (ReferenceDropPanel.Size == ReferenceDropPanel.MaximumSize)
                {
                    ReferencesTimer.Stop();
                    droped = true;
                }
            }

            // if the panel is droped to become smaller with 10 pixels
            else
            {
                ReferenceDropPanel.Height -= 10;

                // if the size of the panes is eaqual to its minimum size to stop the timer and report it as "Not droped"
                if (ReferenceDropPanel.Size == ReferenceDropPanel.MinimumSize)
                {
                    ReferencesTimer.Stop();
                    droped = false;
                }
            }
        }

        #endregion

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            
            if (managerFirstNameTB.Text == "" || managerLastNameTB.Text == "" || managerMiddleNameTB.Text == "")
            {
                warningLabel.Text = "Invalid data";
                warningLabel.Visible = true;
                warningLabel.ForeColor = Color.Red;
            }
            else
            {
                string fullname = managerFirstNameTB.Text + managerMiddleNameTB.Text + managerLastNameTB.Text;
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.InsertInfoDepartments(fullname);
                warningLabel.Text = "Data inserted succeessfuly";
                warningLabel.Visible = true;
                UpdateButton.Visible = true;
                UpdateButton.Enabled = true;
                warningLabel.ForeColor = Color.Green;
            }
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            if (managerFirstNameTB.Text==""||managerLastNameTB.Text==""||managerMiddleNameTB.Text=="")
            {
                warningLabel.ForeColor = Color.Red;
                warningLabel.Text = "Invalid data";
                warningLabel.Visible = true;
            }
            else
            {
                string fullname = managerFirstNameTB.Text + managerMiddleNameTB.Text + managerLastNameTB.Text; 
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.UpdateInfoDepartments(fullname);
                warningLabel.Text = "Data inserted succeessfuly";
                warningLabel.ForeColor = Color.Green;
                warningLabel.Visible = true;
            }
        }

        private void deleteMemberButton_Click(object sender, EventArgs e)
        {
            if (FirstNameTextBox.Text == "" || SecondNameTextBox.Text == "" || LastNameTextBox.Text == "")
            {
                warningLabel.Text = "Invalid data";
                warningLabel.ForeColor = Color.Red;
                warningLabel.Visible = true;
            }
            else
            {
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.DeleteInfoMembers(FirstNameTextBox.Text, SecondNameTextBox.Text, LastNameTextBox.Text);
                warningLabel.Text = "Data deleted succeessfuly";
                warningLabel.ForeColor = Color.Green;
                warningLabel.Visible = true;
                FirstNameTextBox.Text = "";
                SecondNameTextBox.Text = "";
                LastNameTextBox.Text = "";
                //gospodin vasilev ne mi dade inforaciq help
                //durjat me v kilera molq pomosht
                //бтв аз съм сингъл уинкифейс
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            if (FirstNameTextBox.Text == "" || SecondNameTextBox.Text == "" || LastNameTextBox.Text == "")
            {
                warningLabel.Text = "Invalid data";
                warningLabel.Visible = true;
                warningLabel.ForeColor = Color.Red;
            }
            else
            {
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.InsertInfoMembers(FirstNameTextBox.Text, SecondNameTextBox.Text, LastNameTextBox.Text);
                string fullName = FirstNameTextBox.Text + " " + SecondNameTextBox.Text + " " + LastNameTextBox.Text;
                warningLabel.Text = "Data inserted succeessfuly";
                warningLabel.Visible = true;
                warningLabel.ForeColor = Color.Green;
                MemberListBox.Items.Add(fullName);
                FirstNameTextBox.Text = "";
                SecondNameTextBox.Text = "";
                LastNameTextBox.Text = "";
                int numberOfMembers = int.Parse(MemberCounterLabel.Text);
                numberOfMembers++;
                MemberCounterLabel.Text = numberOfMembers.ToString();
            }
        }

        private void RemoveButton_Click(object sender, EventArgs e)
        {
            if (MemberListBox.SelectedItem == null)
            {
                warningLabel.Text = "Please select a name";
                warningLabel.ForeColor = Color.Red;
                warningLabel.Visible = true;
            }
            else
            {
                string fullName = MemberListBox.SelectedItem.ToString();
                string[] names = new string[3];
                names = fullName.Split(' ').ToArray();
                Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
                dataAccess.DeleteInfoMembers(names[0], names[1], names[2]);
                warningLabel.Visible = true;
                int numberOfMembers = int.Parse(MemberCounterLabel.Text);
                numberOfMembers--;
                MemberCounterLabel.Text = numberOfMembers.ToString();
                MemberListBox.Items.Remove(MemberListBox.SelectedItem);
                warningLabel.Text = "Data deleted succeessfuly";
                warningLabel.ForeColor = Color.Green;
            }
        }
    }
}
