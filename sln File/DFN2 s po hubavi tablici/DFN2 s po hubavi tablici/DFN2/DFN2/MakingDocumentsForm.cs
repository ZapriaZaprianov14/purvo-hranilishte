using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Diagnostics;
using System.Data.SqlClient;
using static System.Environment;
using System.Runtime.InteropServices;
namespace DFN2
{
    public partial class MakingDocumentsForm : Form
    {
        Microsoft.Office.Interop.Word.Document wordDoc { get; set; }
        public MakingDocumentsForm()
        {
            InitializeComponent();

        }

        private void MakingDocumentsForm_FormClosed(object sender, FormClosedEventArgs e)
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

        private void CommissionInfoButton_Click(object sender, EventArgs e)
        {
            CommissionInfoForm commissionInfoForm = new CommissionInfoForm();
            this.Hide();
            commissionInfoForm.ShowDialog();
        }

        private void DuplicatesButton_Click(object sender, EventArgs e)
        {
            DuplicateTimer.Start();
        }

        #region Duplicate Number Button
        private void Dup1Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup1Button.Text);
            enterDublicatesForm.Show();
        }

        private void Dup2Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup2Button.Text);
            enterDublicatesForm.Show();
        }

        private void Dup3Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup3Button.Text);
            enterDublicatesForm.Show();
        }

        private void Dup4Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup4Button.Text);
            enterDublicatesForm.Show();
        }

        private void Dup5Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup5Button.Text);
            enterDublicatesForm.Show();
        }

        private void Dup6Button_Click(object sender, EventArgs e)
        {
            EnterDublicatesForm enterDublicatesForm = new EnterDublicatesForm(Dup6Button.Text);
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
            ReferencesForm referencesForm = new ReferencesForm(ReceivedByRequestButton.Text);
            referencesForm.Show();
        }

        private void RegBookButton_Click(object sender, EventArgs e)
        {
            ReferencesForm referencesForm = new ReferencesForm("Издадени по регистрационна книга");
            referencesForm.Show();
        }
        #endregion
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

        #region Reference Timer Function - using it like counter for reference drop menu

        // using this variable to check if the menu is droped or not
        private bool droped = false;
        private void timer1_Tick(object sender, EventArgs e)
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

        private void MakeButton_Click(object sender, EventArgs e)
        {
            string fileName1 = "OtchetKod1.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;

            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitleOtchet1 = Doc.AddTable(1, 12);
            TableTitleOtchet1.Alignment = Alignment.center;
            TableTitleOtchet1.Design = TableDesign.TableGrid;
            TableTitleOtchet1.AutoFit = AutoFit.Contents;
            TableTitleOtchet1.Rows[0].MergeCells(1, 2);
            TableTitleOtchet1.Rows[0].MergeCells(2, 5);
            TableTitleOtchet1.Rows[0].MergeCells(4, 5);
            TableTitleOtchet1.Rows[0].Cells[1].Paragraphs.First().Append("Заявка", FormatTitle);
            TableTitleOtchet1.Rows[0].Cells[2].Paragraphs.First().Append("Получени", FormatTitle);
            TableTitleOtchet1.Rows[0].Cells[4].Paragraphs.First().Append("За унищожаване", FormatTitle);



            Doc.InsertTable(TableTitleOtchet1);

            Table TableTitleOtchet = Doc.AddTable(1, 12);
            TableTitleOtchet.Alignment = Alignment.center;
            TableTitleOtchet.Design = TableDesign.TableGrid;
            TableTitleOtchet.AutoFit = AutoFit.Contents;
            TableTitleOtchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[2].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[3].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[4].Paragraphs.First().Append("от №", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[5].Paragraphs.First().Append("до №", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[6].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[7].Paragraphs.First().Append("Издадени по рег. книга", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[8].Paragraphs.First().Append("Анулирани", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[9].Paragraphs.First().Append("Годни", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[10].Paragraphs.First().Append("Общ брой унищожени", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[11].Paragraphs.First().Append("Остатък дубликати", FormatTitle);



            Doc.InsertTable(TableTitleOtchet);
            //Ном. номер
            Table TableBodyOtchet = Doc.AddTable(14, 12);
            TableBodyOtchet.Alignment = Alignment.center;
            TableBodyOtchet.Design = TableDesign.TableGrid;
            TableBodyOtchet.AutoFit = AutoFit.Contents;
            TableBodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            TableBodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            TableBodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            TableBodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            TableBodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            TableBodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            TableBodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            TableBodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            TableBodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            TableBodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            TableBodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            TableBodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            TableBodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            TableBodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            TableBodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            TableBodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            TableBodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за минали години)");
            TableBodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            TableBodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            TableBodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            TableBodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            TableBodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            TableBodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квал.");
            TableBodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            TableBodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образованието");
            TableBodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            TableBodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            TableBodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");

            //Брой
            TableBodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc1);
            TableBodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc2);
            TableBodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc3);
            TableBodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc4);
            TableBodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc5);
            TableBodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc6);
            TableBodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc7);
            TableBodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc8);
            TableBodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc9);
            TableBodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc10);
            TableBodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc11);
            TableBodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc12);
            TableBodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc13);
            TableBodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc14);
            // Серия
            TableBodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append("C-19");
            TableBodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append("ДС");
            TableBodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append("ОМ-19");
            TableBodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append("ДО");
            TableBodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append("Г-19");
            TableBodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append("ДГ");
            TableBodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append("П-19");
            TableBodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append("ДП");
            TableBodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append("В-19");
            TableBodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append("ДВ");
            TableBodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append("К-19");
            TableBodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append("ДК");
            TableBodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append("СМ-19");
            TableBodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append("О-19");

            //от №
            TableBodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc1Start);
            TableBodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc2Start);
            TableBodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc3Start);
            TableBodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc4Start);
            TableBodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc5Start);
            TableBodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc6Start);
            TableBodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc7Start);
            TableBodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc8Start);
            TableBodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc9Start);
            TableBodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc10Start);
            TableBodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc11Start);
            TableBodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc12Start);
            TableBodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc13Start);
            TableBodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc14Start);


            //до № 
            TableBodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc1End);
            TableBodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc2End);
            TableBodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc3End);
            TableBodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc4End);
            TableBodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc5End);
            TableBodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc6End);
            TableBodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc7End);
            TableBodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc8End);
            TableBodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc9End);
            TableBodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc10End);
            TableBodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc11End);
            TableBodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc12End);
            TableBodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc13End);
            TableBodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc14End);

            //Брой
            TableBodyOtchet.Rows[0].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc1);
            TableBodyOtchet.Rows[1].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc2);
            TableBodyOtchet.Rows[2].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc3);
            TableBodyOtchet.Rows[3].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc4);
            TableBodyOtchet.Rows[4].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc5);
            TableBodyOtchet.Rows[5].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc6);
            TableBodyOtchet.Rows[6].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc7);
            TableBodyOtchet.Rows[7].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc8);
            TableBodyOtchet.Rows[8].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc9);
            TableBodyOtchet.Rows[9].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc10);
            TableBodyOtchet.Rows[10].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc11);
            TableBodyOtchet.Rows[11].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc12);
            TableBodyOtchet.Rows[12].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc13);
            TableBodyOtchet.Rows[13].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc14);

            //Издадени по рег. книга
            TableBodyOtchet.Rows[0].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc1);
            TableBodyOtchet.Rows[1].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc2);
            TableBodyOtchet.Rows[2].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc3);
            TableBodyOtchet.Rows[3].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc4);
            TableBodyOtchet.Rows[4].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc5);
            TableBodyOtchet.Rows[5].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc6);
            TableBodyOtchet.Rows[6].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc7);
            TableBodyOtchet.Rows[7].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc8);
            TableBodyOtchet.Rows[8].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc9);
            TableBodyOtchet.Rows[9].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc10);
            TableBodyOtchet.Rows[10].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc11);
            TableBodyOtchet.Rows[11].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc12);
            TableBodyOtchet.Rows[12].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc13);
            TableBodyOtchet.Rows[13].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc14);

            //Анулирани
            TableBodyOtchet.Rows[0].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc1);
            TableBodyOtchet.Rows[1].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc2);
            TableBodyOtchet.Rows[2].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc3);
            TableBodyOtchet.Rows[3].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc4);
            TableBodyOtchet.Rows[4].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc5);
            TableBodyOtchet.Rows[5].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc6);
            TableBodyOtchet.Rows[6].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc7);
            TableBodyOtchet.Rows[7].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc8);
            TableBodyOtchet.Rows[8].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc9);
            TableBodyOtchet.Rows[9].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc10);
            TableBodyOtchet.Rows[10].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc11);
            TableBodyOtchet.Rows[11].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc12);
            TableBodyOtchet.Rows[12].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc13);
            TableBodyOtchet.Rows[13].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc14);


            //Годни
            TableBodyOtchet.Rows[0].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc1);
            TableBodyOtchet.Rows[1].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc2);
            TableBodyOtchet.Rows[2].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc3);
            TableBodyOtchet.Rows[3].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc4);
            TableBodyOtchet.Rows[4].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc5);
            TableBodyOtchet.Rows[5].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc6);
            TableBodyOtchet.Rows[6].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc7);
            TableBodyOtchet.Rows[7].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc8);
            TableBodyOtchet.Rows[8].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc9);
            TableBodyOtchet.Rows[9].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc10);
            TableBodyOtchet.Rows[10].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc11);
            TableBodyOtchet.Rows[11].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc12);
            TableBodyOtchet.Rows[12].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc13);
            TableBodyOtchet.Rows[13].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc14);

            //Общ брой унищожени
            TableBodyOtchet.Rows[0].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc1Destroyed);
            TableBodyOtchet.Rows[1].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc2Destroyed);
            TableBodyOtchet.Rows[2].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc3Destroyed);
            TableBodyOtchet.Rows[3].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc4Destroyed);
            TableBodyOtchet.Rows[4].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc5Destroyed);
            TableBodyOtchet.Rows[5].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc6Destroyed);
            TableBodyOtchet.Rows[6].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc7Destroyed);
            TableBodyOtchet.Rows[7].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc8Destroyed);
            TableBodyOtchet.Rows[8].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc9Destroyed);
            TableBodyOtchet.Rows[9].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc10Destroyed);
            TableBodyOtchet.Rows[10].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc11Destroyed);
            TableBodyOtchet.Rows[11].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc12Destroyed);
            TableBodyOtchet.Rows[12].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc13Destroyed);
            TableBodyOtchet.Rows[13].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc14Destroyed);

            //Остатък дубликати
            TableBodyOtchet.Rows[0].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc1Rest);
            TableBodyOtchet.Rows[1].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc2Rest);
            TableBodyOtchet.Rows[2].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc3Rest);
            TableBodyOtchet.Rows[3].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc4Rest);
            TableBodyOtchet.Rows[4].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc5Rest);
            TableBodyOtchet.Rows[5].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc6Rest);
            TableBodyOtchet.Rows[6].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc7Rest);
            TableBodyOtchet.Rows[7].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc8Rest);
            TableBodyOtchet.Rows[8].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc9Rest);
            TableBodyOtchet.Rows[9].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc10Rest);
            TableBodyOtchet.Rows[10].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc11Rest);
            TableBodyOtchet.Rows[11].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc12Rest);
            TableBodyOtchet.Rows[12].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc13Rest);
            TableBodyOtchet.Rows[13].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc14Rest);

            Doc.InsertTable(TableBodyOtchet);

            Doc.InsertParagraphs("");
            
            //табле 2


            Formatting FormatTitle1 = new Formatting();
            FormatTitle1.Size = 11D;
            FormatTitle1.Bold = true;
            FormatTitle1.Position = 20;
            FormatTitle1.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitleOtchet2 = Doc.AddTable(1, 11);
            TableTitleOtchet2.Alignment = Alignment.center;
            TableTitleOtchet2.Design = TableDesign.TableGrid;
            TableTitleOtchet2.AutoFit = AutoFit.Contents;
            TableTitleOtchet2.Rows[0].MergeCells(2, 5);
            TableTitleOtchet2.Rows[0].MergeCells(4, 5);
            TableTitleOtchet2.Rows[0].Cells[2].Paragraphs.First().Append("Получени от други институции", FormatTitle);
            TableTitleOtchet2.Rows[0].Cells[4].Paragraphs.First().Append("За унищожаване", FormatTitle);


            Doc.InsertTable(TableTitleOtchet2);

            Table TableTitle2Otchet = Doc.AddTable(1, 11);
            TableTitle2Otchet.Alignment = Alignment.center;
            TableTitle2Otchet.Design = TableDesign.TableGrid;
            TableTitle2Otchet.AutoFit = AutoFit.Contents;
            TableTitle2Otchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[2].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[3].Paragraphs.First().Append("от №", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[4].Paragraphs.First().Append("до №", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[5].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[6].Paragraphs.First().Append("Издадени по рег. книга", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[7].Paragraphs.First().Append("Анулирани", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[8].Paragraphs.First().Append("Годни", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[9].Paragraphs.First().Append("Общ брой унищожени", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[10].Paragraphs.First().Append("Остатък дубликати", FormatTitle);



            Doc.InsertTable(TableTitle2Otchet);

            //Ном. номер
            Table Table2BodyOtchet = Doc.AddTable(14, 11);
            TableBodyOtchet.Alignment = Alignment.center;
            TableBodyOtchet.Design = TableDesign.TableGrid;
            Table2BodyOtchet.AutoFit = AutoFit.Fixed;
            Table2BodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            Table2BodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            Table2BodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            Table2BodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            Table2BodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            Table2BodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            Table2BodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            Table2BodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            Table2BodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            Table2BodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            Table2BodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            Table2BodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            Table2BodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            Table2BodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            Table2BodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            Table2BodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            Table2BodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за минали години)");
            Table2BodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table2BodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table2BodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            Table2BodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            Table2BodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            Table2BodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квалификация");
            Table2BodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table2BodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален етап");
            Table2BodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table2BodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            Table2BodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");


            // Серия
            Table2BodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append("C-19");
            Table2BodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append("ДС");
            Table2BodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append("ОМ-19");
            Table2BodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append("ДО");
            Table2BodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append("Г-19");
            Table2BodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append("ДГ");
            Table2BodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append("П-19");
            Table2BodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append("ДП");
            Table2BodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append("В-19");
            Table2BodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append("ДВ");
            Table2BodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append("К-19");
            Table2BodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append("ДК");
            Table2BodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append("СМ-19");
            Table2BodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append("О-19");


            //от №
            Table2BodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc1Start);
            Table2BodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc2Start);
            Table2BodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc3Start);
            Table2BodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc4Start);
            Table2BodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc5Start);
            Table2BodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc6Start);
            Table2BodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc7Start);
            Table2BodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc8Start);
            Table2BodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc9Start);
            Table2BodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc10Start);
            Table2BodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc11Start);
            Table2BodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc12Start);
            Table2BodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc13Start);
            Table2BodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc14Start);

            //до №
            Table2BodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc1End);
            Table2BodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc2End);
            Table2BodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc3End);
            Table2BodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc4End);
            Table2BodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc5End);
            Table2BodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc6End);
            Table2BodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc7End);
            Table2BodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc8End);
            Table2BodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc9End);
            Table2BodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc10End);
            Table2BodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc11End);
            Table2BodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc12End);
            Table2BodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc13End);
            Table2BodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc14End);


            //Брой
            Table2BodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc1);
            Table2BodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc2);
            Table2BodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc3);
            Table2BodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc4);
            Table2BodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc5);
            Table2BodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc6);
            Table2BodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc7);
            Table2BodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc8);
            Table2BodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc9);
            Table2BodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc10);
            Table2BodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc11);
            Table2BodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc12);
            Table2BodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc13);
            Table2BodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc14);


            //Издадени по рег. книга
            Table2BodyOtchet.Rows[0].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc1);
            Table2BodyOtchet.Rows[1].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc2);
            Table2BodyOtchet.Rows[2].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc3);
            Table2BodyOtchet.Rows[3].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc4);
            Table2BodyOtchet.Rows[4].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc5);
            Table2BodyOtchet.Rows[5].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc6);
            Table2BodyOtchet.Rows[6].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc7);
            Table2BodyOtchet.Rows[7].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc8);
            Table2BodyOtchet.Rows[8].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc9);
            Table2BodyOtchet.Rows[9].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc10);
            Table2BodyOtchet.Rows[10].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc11);
            Table2BodyOtchet.Rows[11].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc12);
            Table2BodyOtchet.Rows[12].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc13);
            Table2BodyOtchet.Rows[13].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc14);


            //Анулирани
            Table2BodyOtchet.Rows[0].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc1);
            Table2BodyOtchet.Rows[1].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc2);
            Table2BodyOtchet.Rows[2].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc3);
            Table2BodyOtchet.Rows[3].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc4);
            Table2BodyOtchet.Rows[4].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc5);
            Table2BodyOtchet.Rows[5].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc6);
            Table2BodyOtchet.Rows[6].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc7);
            Table2BodyOtchet.Rows[7].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc8);
            Table2BodyOtchet.Rows[8].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc9);
            Table2BodyOtchet.Rows[9].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc10);
            Table2BodyOtchet.Rows[10].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc11);
            Table2BodyOtchet.Rows[11].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc12);
            Table2BodyOtchet.Rows[12].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc13);
            Table2BodyOtchet.Rows[13].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc14);


            //Годни
            Table2BodyOtchet.Rows[0].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc1);
            Table2BodyOtchet.Rows[1].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc2);
            Table2BodyOtchet.Rows[2].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc3);
            Table2BodyOtchet.Rows[3].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc4);
            Table2BodyOtchet.Rows[4].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc5);
            Table2BodyOtchet.Rows[5].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc6);
            Table2BodyOtchet.Rows[6].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc7);
            Table2BodyOtchet.Rows[7].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc8);
            Table2BodyOtchet.Rows[8].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc9);
            Table2BodyOtchet.Rows[9].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc10);
            Table2BodyOtchet.Rows[10].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc11);
            Table2BodyOtchet.Rows[11].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc12);
            Table2BodyOtchet.Rows[12].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc13);
            Table2BodyOtchet.Rows[13].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc14);

            //Общ брой унищожени
            Table2BodyOtchet.Rows[0].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc1Destroyed);
            Table2BodyOtchet.Rows[1].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc2Destroyed);
            Table2BodyOtchet.Rows[2].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc3Destroyed);
            Table2BodyOtchet.Rows[3].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc4Destroyed);
            Table2BodyOtchet.Rows[4].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc5Destroyed);
            Table2BodyOtchet.Rows[5].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc6Destroyed);
            Table2BodyOtchet.Rows[6].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc7Destroyed);
            Table2BodyOtchet.Rows[7].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc8Destroyed);
            Table2BodyOtchet.Rows[8].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc9Destroyed);
            Table2BodyOtchet.Rows[9].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc10Destroyed);
            Table2BodyOtchet.Rows[10].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc11Destroyed);
            Table2BodyOtchet.Rows[11].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc12Destroyed);
            Table2BodyOtchet.Rows[12].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc13Destroyed);
            Table2BodyOtchet.Rows[13].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc14Destroyed);


            //Остатък дубликати
            Table2BodyOtchet.Rows[0].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc1Rest);
            Table2BodyOtchet.Rows[1].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc2Rest);
            Table2BodyOtchet.Rows[2].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc3Rest);
            Table2BodyOtchet.Rows[3].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc4Rest);
            Table2BodyOtchet.Rows[4].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc5Rest);
            Table2BodyOtchet.Rows[5].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc6Rest);
            Table2BodyOtchet.Rows[6].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc7Rest);
            Table2BodyOtchet.Rows[7].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc8Rest);
            Table2BodyOtchet.Rows[8].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc9Rest);
            Table2BodyOtchet.Rows[9].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc10Rest);
            Table2BodyOtchet.Rows[10].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc11Rest);
            Table2BodyOtchet.Rows[11].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc12Rest);
            Table2BodyOtchet.Rows[12].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc13Rest);
            Table2BodyOtchet.Rows[13].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc14Rest);



            Doc.InsertTable(Table2BodyOtchet);

            Doc.InsertParagraphs("");

            //table 3
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;
            Formatting FormatTitle3 = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitleOtchet3 = Doc.AddTable(1, 6);
            TableTitleOtchet3.Alignment = Alignment.center;
            TableTitleOtchet3.Design = TableDesign.TableGrid;
            TableTitleOtchet3.AutoFit = AutoFit.Contents;
            TableTitleOtchet3.Rows[0].MergeCells(2, 6);
            TableTitleOtchet3.Rows[0].Cells[2].Paragraphs.First().Append("Предадени на други институции", FormatTitle);


            Doc.InsertTable(TableTitleOtchet3);

            Table Table3TitleOtchet = Doc.AddTable(1, 6);
            Table3TitleOtchet.Alignment = Alignment.center;
            Table3TitleOtchet.Design = TableDesign.TableGrid;
            Table3TitleOtchet.AutoFit = AutoFit.Contents;
            Table3TitleOtchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[2].Paragraphs.First().Append("Серия", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[3].Paragraphs.First().Append("от №", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[4].Paragraphs.First().Append("до №", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[5].Paragraphs.First().Append("Брой", FormatTitle);



            Doc.InsertTable(Table3TitleOtchet);


            Table Table3BodyOtchet = Doc.AddTable(14, 6);
            Table3BodyOtchet.Alignment = Alignment.center;
            Table3BodyOtchet.Design = TableDesign.TableGrid;
            Table3BodyOtchet.AutoFit = AutoFit.Fixed;
            Table3BodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            Table3BodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            Table3BodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            Table3BodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            Table3BodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            Table3BodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            Table3BodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            Table3BodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            Table3BodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            Table3BodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            Table3BodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            Table3BodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            Table3BodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            Table3BodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            Table3BodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            Table3BodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            Table3BodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за мин. год.)");
            Table3BodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table3BodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table3BodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            Table3BodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            Table3BodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            Table3BodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квалификация");
            Table3BodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table3BodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образованието");
            Table3BodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table3BodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            Table3BodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");
            // Серия
            Table3BodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append("C-19");
            Table3BodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append("ДС");
            Table3BodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append("ОМ-19");
            Table3BodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append("ДО");
            Table3BodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append("Г-19");
            Table3BodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append("ДГ");
            Table3BodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append("П-19");
            Table3BodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append("ДП");
            Table3BodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append("В-19");
            Table3BodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append("ДВ");
            Table3BodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append("К-19");
            Table3BodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append("ДК");
            Table3BodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append("СМ-19");
            Table3BodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append("О-19");

            //от №
            Table3BodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc1Start);
            Table3BodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc2Start);
            Table3BodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc3Start);
            Table3BodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc4Start);
            Table3BodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc5Start);
            Table3BodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc6Start);
            Table3BodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc7Start);
            Table3BodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc8Start);
            Table3BodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc9Start);
            Table3BodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc10Start);
            Table3BodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc11Start);
            Table3BodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc12Start);
            Table3BodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc13Start);
            Table3BodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc14Start);


            //до № 
            Table3BodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc1End);
            Table3BodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc2End);
            Table3BodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc3End);
            Table3BodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc4End);
            Table3BodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc5End);
            Table3BodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc6End);
            Table3BodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc7End);
            Table3BodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc8End);
            Table3BodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc9End);
            Table3BodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc10End);
            Table3BodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc11End);
            Table3BodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc12End);
            Table3BodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc13End);
            Table3BodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc14End);


            //Брой
            Table3BodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc1Count);
            Table3BodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc2Count);
            Table3BodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc3Count);
            Table3BodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc4Count);
            Table3BodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc5Count);
            Table3BodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc6Count);
            Table3BodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc7Count);
            Table3BodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc8Count);
            Table3BodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc9Count);
            Table3BodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc10Count);
            Table3BodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc11Count);
            Table3BodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc12Count);
            Table3BodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc13Count);
            Table3BodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc14Count);





            Doc.InsertTable(Table3BodyOtchet);

            Doc.InsertParagraphs("");
            //table 4

            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;


            Formatting FormatTitle4 = new Formatting();
            FormatTitle4.Size = 11D;
            FormatTitle4.Bold = true;
            FormatTitle4.Position = 20;
            FormatTitle4.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitle4 = Doc.AddTable(1, 6);
            TableTitle4.Design = TableDesign.TableGrid;
            TableTitle4.AutoFit = AutoFit.Contents;
            TableTitle4.Rows[0].Cells[0].Paragraphs.First().Append("Година на получаване", FormatTitle);
            TableTitle4.Rows[0].Cells[1].Paragraphs.First().Append("Номенклатурен №", FormatTitle);
            TableTitle4.Rows[0].Cells[2].Paragraphs.First().Append("Наименование на дубликата", FormatTitle);
            TableTitle4.Rows[0].Cells[3].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitle4.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitle4.Rows[0].Cells[5].Paragraphs.First().Append("Описание на фабричните номера на наличните дубликати", FormatTitle);

            Doc.InsertTable(TableTitle4);
            Table Table4 = Doc.AddTable(8, 6);
            //Година на получаване
            Table4.Rows[0].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[1].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[2].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[3].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[4].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[5].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[6].Cells[0].Paragraphs.First().Append("2019");
            Table4.Rows[7].Cells[0].Paragraphs.First().Append("2019");


            //ном.номер
            Table4.Rows[0].Cells[1].Paragraphs.First().Append("3-44а");
            Table4.Rows[1].Cells[1].Paragraphs.First().Append("3-30а");
            Table4.Rows[2].Cells[1].Paragraphs.First().Append("3-22а");
            Table4.Rows[3].Cells[1].Paragraphs.First().Append("3-54а");
            Table4.Rows[4].Cells[1].Paragraphs.First().Append("3-54аВ");
            Table4.Rows[5].Cells[1].Paragraphs.First().Append("3-27аВ");
            Table4.Rows[6].Cells[1].Paragraphs.First().Append("3-34а");
            Table4.Rows[7].Cells[1].Paragraphs.First().Append("3-30а");


            //наименование на дубл.
            Table4.Rows[0].Cells[2].Paragraphs.First().Append("Дубликат на диплома");
            Table4.Rows[1].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table4.Rows[2].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            Table4.Rows[3].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за проф. квалификация");
            Table4.Rows[4].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table4.Rows[5].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за валид. на комп.");
            Table4.Rows[6].Cells[2].Paragraphs.First().Append("Дубликат на диплома");
            Table4.Rows[7].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");


            //брой
            Table4.Rows[0].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[1].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[2].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[3].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[4].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[5].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[6].Cells[3].Paragraphs.First().Append("");
            Table4.Rows[7].Cells[3].Paragraphs.First().Append("");


            //серия
            Table4.Rows[0].Cells[4].Paragraphs.First().Append("ДС");
            Table4.Rows[1].Cells[4].Paragraphs.First().Append("ДО");
            Table4.Rows[2].Cells[4].Paragraphs.First().Append("ДГ");
            Table4.Rows[3].Cells[4].Paragraphs.First().Append("ДП");
            Table4.Rows[4].Cells[4].Paragraphs.First().Append("ДВ");
            Table4.Rows[5].Cells[4].Paragraphs.First().Append("ДК");
            Table4.Rows[6].Cells[4].Paragraphs.First().Append("ДС");
            Table4.Rows[7].Cells[4].Paragraphs.First().Append("ДО");


            //остатък
            Table4.Rows[0].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[1].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[2].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[3].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[4].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[5].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[6].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[7].Cells[5].Paragraphs.First().Append("");




            Doc.InsertTable(Table4);

            Table Table4a = Doc.AddTable(4, 6);
            //година
            Table4a.Rows[0].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[1].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[2].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[2].Cells[0].Paragraphs.First().Append("2019");
            //ном.номер
            Table4a.Rows[0].Cells[1].Paragraphs.First().Append("3-54а");
            Table4a.Rows[1].Cells[1].Paragraphs.First().Append("3-54аВ");
            Table4a.Rows[2].Cells[1].Paragraphs.First().Append("3-27аВ");
            Table4a.Rows[3].Cells[1].Paragraphs.First().Append("3-22а");
            //наименование
            Table4a.Rows[0].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за проф. квалификация");
            Table4a.Rows[1].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table4a.Rows[2].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table4a.Rows[3].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            //брой
            Table4a.Rows[0].Cells[3].Paragraphs.First().Append("");
            Table4a.Rows[1].Cells[3].Paragraphs.First().Append("");
            Table4a.Rows[2].Cells[3].Paragraphs.First().Append("");
            Table4a.Rows[3].Cells[3].Paragraphs.First().Append("");
            //серия
            Table4a.Rows[0].Cells[4].Paragraphs.First().Append("ДП");
            Table4a.Rows[1].Cells[4].Paragraphs.First().Append("ДВ");
            Table4a.Rows[2].Cells[4].Paragraphs.First().Append("ДК");
            Table4a.Rows[3].Cells[4].Paragraphs.First().Append("ДГ");
            //остатък
            Table4a.Rows[0].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[1].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[2].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[3].Cells[5].Paragraphs.First().Append("");

            Doc.InsertTable(Table4a);
            Process.Start("WINWORD.EXE", fileName1);

            Doc.Save();
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName1 = "Zaqvka.docx";
            var Doc = DocX.Create(fileName1);
            string Title1 = " МИНИСТЕРСТВО  НА ОБРАЗОВАНИЕТО  И НАУКАТА";
            Formatting FormatTitle1 = new Formatting();
            FormatTitle1.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            FormatTitle1.Bold = true;
            FormatTitle1.Size = 14D;
            FormatTitle1.Position = 40;
            FormatTitle1.UnderlineColor = Color.Black;
            Paragraph Titleparagraph = Doc.InsertParagraph(Title1, false, FormatTitle1);
            Titleparagraph.Alignment = Alignment.center;

            string Title2 = "РЕГИОНАЛНО УПРАВЛЕНИЕ НА ОБРАЗОВАНИЕТО";
            Formatting FormatTitle2 = new Formatting();
            FormatTitle2.Size = 12D;
            FormatTitle2.Bold = true;
            FormatTitle2.Position = 40;
            FormatTitle2.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph2 = Doc.InsertParagraph(Title2, false, FormatTitle2);
            Titleparagraph2.Alignment = Alignment.center;

            string Title3 = "……………………………………………";
            Paragraph Titleparagraph3 = Doc.InsertParagraph(Title3, false, FormatTitle2);
            Titleparagraph3.Alignment = Alignment.center;
            string Title4 = "З  А  Я  В  К  А";
            Formatting FormatTitle4 = new Formatting();
            FormatTitle4.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            FormatTitle4.Bold = true;
            FormatTitle4.Size = 14D;
            FormatTitle4.Position = 40;
            Paragraph Titleparagraph4 = Doc.InsertParagraph(Title4, false, FormatTitle4);
            Titleparagraph4.Alignment = Alignment.center;
            string Title5 = "за доставка на документи от задължителната документация за системата на" +
                " предучилищното и училищното образование за края на учебната 2019/2020 година";
            Paragraph Titleparagraph5 = Doc.InsertParagraph(Title5, false, FormatTitle2);
            Titleparagraph5.Alignment = Alignment.center;
            string Title6 = "……………………………………………………………………………………………";
            Formatting FormatTitle6 = new Formatting();
            FormatTitle6.Size = 12D;
            FormatTitle6.Bold = false;
            FormatTitle6.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph6 = Doc.InsertParagraph(Title6, false, FormatTitle2);
            Titleparagraph6.Alignment = Alignment.center;
            string Title7 = "трите имена на директора/началника на РУО ";
            Formatting FormatTitle7 = new Formatting();
            FormatTitle7.Size = 9D;
            FormatTitle7.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph7 = Doc.InsertParagraph(Title7, false, FormatTitle7);
            Titleparagraph7.Alignment = Alignment.center;
            string Title8 = "………………………………………………………………………………………………";
            Paragraph Titleparagraph8 = Doc.InsertParagraph(Title8, false, FormatTitle2);
            Titleparagraph8.Alignment = Alignment.center;
            string Title9 = "пълно наименование на институцията";
            Paragraph Titleparagraph9 = Doc.InsertParagraph(Title9, false, FormatTitle7);
            Titleparagraph9.Alignment = Alignment.center;
            string TitlePrazenRed = "";
            Paragraph Titleparagraphprazen = Doc.InsertParagraph(TitlePrazenRed, false, FormatTitle6);
            Titleparagraphprazen.Alignment = Alignment.center;
            string Title10 = "телефон: …………………………" + " e-mail:……………………………………";
            Paragraph Titleparagraph10 = Doc.InsertParagraph(Title10, false, FormatTitle2);
            Titleparagraph10.Alignment = Alignment.left;
            string Title11 = "адрес на институцията:…………………………………………………………………";
            Paragraph Titleparagraph11 = Doc.InsertParagraph(Title11, false, FormatTitle6);
            Titleparagraph11.Alignment = Alignment.left;
            string Title12 = "";
            Paragraph Titleparagraph12 = Doc.InsertParagraph(Title12, false, FormatTitle6);
            Titleparagraph12.Alignment = Alignment.center;
            string TableTitl1 = "Номер";
            string TableTitl2 = "Наименование на документа";
            string TableTitl3 = "Заявено количество(брой)";

            Table Table1 = Doc.AddTable(13, 3);
            Table1.Alignment = Alignment.center;
            Table1.Design = TableDesign.TableGrid;
            Table1.AutoFit = AutoFit.Window;
            Table1.AutoFit = AutoFit.Contents;
            Table1.Rows[0].Cells[0].Paragraphs.First().Append(TableTitl1);
            Table1.Rows[1].Cells[0].Paragraphs.First().Append("3-19");
            Table1.Rows[2].Cells[0].Paragraphs.First().Append("3-23");
            Table1.Rows[3].Cells[0].Paragraphs.First().Append("3-25");
            Table1.Rows[4].Cells[0].Paragraphs.First().Append("3-20");
            Table1.Rows[5].Cells[0].Paragraphs.First().Append("3-30");
            Table1.Rows[6].Cells[0].Paragraphs.First().Append("3-30а");
            Table1.Rows[7].Cells[0].Paragraphs.First().Append("3-22");
            Table1.Rows[8].Cells[0].Paragraphs.First().Append("3-22а");
            Table1.Rows[9].Cells[0].Paragraphs.First().Append("3-22.1");
            Table1.Rows[10].Cells[0].Paragraphs.First().Append("3-34");
            Table1.Rows[11].Cells[0].Paragraphs.First().Append("3-44а");
            Table1.Rows[12].Cells[0].Paragraphs.First().Append("3-34.1");

            Table1.Rows[0].Cells[1].Paragraphs.First().Append(TableTitl2);
            Table1.Rows[1].Cells[1].Paragraphs.First().Append("Удостоверение за задължително предучилищно образование");
            Table1.Rows[2].Cells[1].Paragraphs.First().Append("Удостоверение за завършен клас от начален етап на основно образование (за I клас, II клас и III клас)");
            Table1.Rows[3].Cells[1].Paragraphs.First().Append("Удостоверение за завършен начален етап на основно образование – IV клас");
            Table1.Rows[4].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование – (образец за минали години)");
            Table1.Rows[5].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");
            Table1.Rows[6].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование ");
            Table1.Rows[7].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table1.Rows[8].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            Table1.Rows[9].Cells[1].Paragraphs.First().Append("Приложение към удостоверение за завършен гимназиален етап ");
            Table1.Rows[10].Cells[1].Paragraphs.First().Append("Диплома за средно образование - за випуск 2019");
            Table1.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на диплома за средно образование ");
            Table1.Rows[12].Cells[1].Paragraphs.First().Append("Приложение към диплома за средно образование ");

            Table1.Rows[0].Cells[2].Paragraphs.First().Append(TableTitl3);


            Doc.InsertTable(Table1);
            Paragraph Titleparagraph101 = Doc.InsertParagraph(Title12, false, FormatTitle6);
            Titleparagraph12.Alignment = Alignment.center;

            Table Table2 = Doc.AddTable(15, 3);
            Table2.Alignment = Alignment.center;
            Table2.Design = TableDesign.TableGrid;
            Table2.AutoFit = AutoFit.Window;
            Table2.AutoFit = AutoFit.Contents;
            Table2.Rows[0].Cells[0].Paragraphs.First().Append(TableTitl1);
            Table2.Rows[1].Cells[0].Paragraphs.First().Append("3-37");
            Table2.Rows[2].Cells[0].Paragraphs.First().Append("3-54");
            Table2.Rows[3].Cells[0].Paragraphs.First().Append("3-54а");
            Table2.Rows[4].Cells[0].Paragraphs.First().Append("3-54.1");
            Table2.Rows[5].Cells[0].Paragraphs.First().Append("3-37В");
            Table2.Rows[6].Cells[0].Paragraphs.First().Append("3-54В");
            Table2.Rows[7].Cells[0].Paragraphs.First().Append("3-54аВ");
            Table2.Rows[8].Cells[0].Paragraphs.First().Append("3-102");
            Table2.Rows[9].Cells[0].Paragraphs.First().Append("3-27В");
            Table2.Rows[10].Cells[0].Paragraphs.First().Append("3-27аВ");
            Table2.Rows[11].Cells[0].Paragraphs.First().Append("3-103");
            Table2.Rows[12].Cells[0].Paragraphs.First().Append("3-114");
            Table2.Rows[13].Cells[0].Paragraphs.First().Append("3-116");
            Table2.Rows[14].Cells[0].Paragraphs.First().Append("3-42");


            Table2.Rows[0].Cells[1].Paragraphs.First().Append(TableTitl2);
            Table2.Rows[1].Cells[1].Paragraphs.First().Append("Удостоверение за професионално обучение");
            Table2.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            Table2.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация ");
            Table2.Rows[4].Cells[1].Paragraphs.First().Append("Приложение към свидетелство за професионална квалификация ");
            Table2.Rows[5].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на професионална квалификация по част от професията");
            Table2.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за валидиране на професионална квалификация");
            Table2.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране на професионална квалификация");
            Table2.Rows[8].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности по учебен предмет, невключен в дипломата за средно образование");
            Table2.Rows[9].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образование ");
            Table2.Rows[10].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образование ");
            Table2.Rows[11].Cells[1].Paragraphs.First().Append("Удостоверение за завършен клас ");
            Table2.Rows[12].Cells[1].Paragraphs.First().Append("Свидетелство за правоспособност ");
            Table2.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за правоспособност по заваряване");
            Table2.Rows[14].Cells[1].Paragraphs.First().Append("Диплома за средно образование (образец за минали години)");

            Table2.Rows[0].Cells[2].Paragraphs.First().Append(TableTitl3);


            Doc.InsertTable(Table2);


            string Title13 = "…………………………………………………………………………………………………";
            Paragraph Titleparagraph13 = Doc.InsertParagraph(Title13, false, FormatTitle6);
            Titleparagraph13.Alignment = Alignment.center;
            string Title14 = "пълно наименование на платеца на задължителната документация";
            Paragraph Titleparagraph14 = Doc.InsertParagraph(Title14, false, FormatTitle7);
            Titleparagraph14.Alignment = Alignment.center;
            string Title15 = "БУЛСТАТ …………………………………………………………….";
            Paragraph Titleparagraph15 = Doc.InsertParagraph(Title15, false, FormatTitle6);
            Titleparagraph15.Alignment = Alignment.left;
            string Title16 = "на платеца";
            Paragraph Titleparagraph16 = Doc.InsertParagraph(Title16, false, FormatTitle7);
            Titleparagraph16.Alignment = Alignment.center;
            string Title17 = "ДИРЕКТОР …………………………………………………         ............................";
            Paragraph Titleparagraph17 = Doc.InsertParagraph(Title17, false, FormatTitle6);
            Titleparagraph17.Alignment = Alignment.left;
            string Title18 = "име и фамилия                                                                                    подпис и печат";
            Paragraph Titleparagraph18 = Doc.InsertParagraph(Title18, false, FormatTitle7);
            Titleparagraph18.Alignment = Alignment.center;
            string Title19a = "ОБЩИНСКА";
            Paragraph Titleparagraph19a = Doc.InsertParagraph(Title19a, false, FormatTitle6);
            Titleparagraph19a.Alignment = Alignment.left;
            string Title19 = " АДМИНИСТРАЦИЯ  ……………………………………   ……………   тел: …………… ";
            Paragraph Titleparagraph19 = Doc.InsertParagraph(Title19, false, FormatTitle6);
            Titleparagraph19.Alignment = Alignment.left;
            string Title20 = "      име, фамилия и длъжност на лицето         подпис и печат";
            Paragraph Titleparagraph20 = Doc.InsertParagraph(Title20, false, FormatTitle7);
            Titleparagraph20.Alignment = Alignment.center;
            string Title21a = "ПРОВЕРИЛ:";
            Paragraph Titleparagraph21a = Doc.InsertParagraph(Title21a, false, FormatTitle6);
            Titleparagraph21a.Alignment = Alignment.left;
            string Title21 = "ЕКСПЕРТ В РУО …………………………………       ………………     тел: …………";
            Paragraph Titleparagraph21 = Doc.InsertParagraph(Title21, false, FormatTitle6);
            Titleparagraph21.Alignment = Alignment.left;
            string Title22 = "     име и фамилия                                                    подпис      ";
            Paragraph Titleparagraph22 = Doc.InsertParagraph(Title22, false, FormatTitle7);
            Titleparagraph22.Alignment = Alignment.center;
            string Title23a = "НАЧАЛНИК";
            Paragraph Titleparagraph23a = Doc.InsertParagraph(Title23a, false, FormatTitle6);
            Titleparagraph23a.Alignment = Alignment.left;
            string Title23 = "НА РУО……………………….............    ................................            тел: …………. ";
            Paragraph Titleparagraph23 = Doc.InsertParagraph(Title23, false, FormatTitle6);
            Titleparagraph23.Alignment = Alignment.left;
            string Title24 = "  име и фамилия 		            подпис и печат				";
            Paragraph Titleparagraph24 = Doc.InsertParagraph(Title24, false, FormatTitle7);
            Titleparagraph24.Alignment = Alignment.center;
            Doc.Save();
            Process.Start("WINWORD.EXE", fileName1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string fileName1 = "OtchetKod2.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;

            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess1 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess2 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess3 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess4 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess5 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess6 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess7 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess8 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess9 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess10 = new Database_Elements.DataAccess();
            string schoolNeispuo = dataAccess.GetSchoolNeispuo();
            string street = dataAccess1.GetSchoolStreet();
            string city = dataAccess2.GetSchoolCity();
            string community = dataAccess3.GetSchoolCommunity();
            string area = dataAccess4.GetSchoolArea();
            string postCode = dataAccess5.GetSchoolPostCode();
            string phone = dataAccess6.GetSchoolPhone();
            string email = dataAccess7.GetSchoolEmail();
            string principalName = dataAccess8.GetSchoolPrincipalName();
            string manager = dataAccess9.GetManager();
            string schoolName = dataAccess10.GetSchoolName();
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph = Doc.InsertParagraph("Приложение №6 към чл.52 ал.1");
            Titleparagraph.Alignment = Alignment.center;
            

            Paragraph Titleparagraph1 = Doc.InsertParagraph("Изм. - ДБ,бр.75 от 2017г.,в сила от 15.09.2017");
            Titleparagraph1.Alignment = Alignment.center;


            Paragraph Titleparagraph2 = Doc.InsertParagraph(schoolName);
            Titleparagraph2.Alignment = Alignment.center;

            Paragraph Titleparagraph3 = Doc.InsertParagraph($"гр.(с){city} {street},община {community},област {area}");
            Titleparagraph3.Alignment = Alignment.center;

            Paragraph Titleparagraph4 = Doc.InsertParagraph("ОТЧЕТ");
            Titleparagraph4.Alignment = Alignment.center;

            Paragraph Titleparagraph5 = Doc.InsertParagraph($"на документите с фабрична номерация");
            Titleparagraph5.Alignment = Alignment.center;

            Paragraph Titleparagraph6 = Doc.InsertParagraph($"Днес, 20.04.2020, комисия, назначена със Заповед № {Globals.OrderNumber} от {Globals.OrderDate.ToString()} на директора в състав:");
            Titleparagraph6.Alignment = Alignment.left;

            Paragraph Titleparagraph7 = Doc.InsertParagraph($"Председател:{manager}");
            Titleparagraph7.Alignment = Alignment.left;

            Paragraph Titleparagraph8 = Doc.InsertParagraph("Членове:");
            Titleparagraph8.Alignment = Alignment.left;

            SqlConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand("SELECT * FROM department_members WHERE id=id", connection);
                SqlDataReader reader = command.ExecuteReader();
                int i = 1;
                using (reader)
                {
                    while (reader.Read())
                    {
                        string fullName = (string)reader["first_name"] + " " + (string)reader["middle_name"] + " " + (string)reader["last_name"];
                        Doc.InsertParagraph($"{i}. {fullName}");
                    }
                }
            }
            connection.Close();
            Doc.InsertParagraph("на заседание отчете документите с фабрична номерация за учебната 2018/2019 години,както следва");
            Table TableTitleOtchet1 = Doc.AddTable(1, 12);
            TableTitleOtchet1.Alignment = Alignment.center;
            TableTitleOtchet1.Design = TableDesign.TableGrid;
            TableTitleOtchet1.AutoFit = AutoFit.Contents;
            TableTitleOtchet1.Rows[0].MergeCells(1, 2);
            TableTitleOtchet1.Rows[0].MergeCells(2, 5);
            TableTitleOtchet1.Rows[0].MergeCells(4, 5);
            TableTitleOtchet1.Rows[0].Cells[1].Paragraphs.First().Append("Заявка", FormatTitle);
            TableTitleOtchet1.Rows[0].Cells[2].Paragraphs.First().Append("Получени", FormatTitle);
            TableTitleOtchet1.Rows[0].Cells[4].Paragraphs.First().Append("За унищожаване", FormatTitle);



            Doc.InsertTable(TableTitleOtchet1);

            Table TableTitleOtchet = Doc.AddTable(1, 12);
            TableTitleOtchet.Alignment = Alignment.center;
            TableTitleOtchet.Design = TableDesign.TableGrid;
            TableTitleOtchet.AutoFit = AutoFit.Contents;
            TableTitleOtchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[2].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[3].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[4].Paragraphs.First().Append("от №", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[5].Paragraphs.First().Append("до №", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[6].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[7].Paragraphs.First().Append("Издадени по рег. книга", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[8].Paragraphs.First().Append("Анулирани", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[9].Paragraphs.First().Append("Годни", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[10].Paragraphs.First().Append("Общ брой унищожени", FormatTitle);
            TableTitleOtchet.Rows[0].Cells[11].Paragraphs.First().Append("Остатък дубликати", FormatTitle);



            Doc.InsertTable(TableTitleOtchet);
            //Ном. номер
            Table TableBodyOtchet = Doc.AddTable(14, 12);
            TableBodyOtchet.Alignment = Alignment.center;
            TableBodyOtchet.Design = TableDesign.TableGrid;
            TableBodyOtchet.AutoFit = AutoFit.Contents;
            TableBodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            TableBodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            TableBodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            TableBodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            TableBodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            TableBodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            TableBodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            TableBodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            TableBodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            TableBodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            TableBodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            TableBodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            TableBodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            TableBodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            TableBodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            TableBodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            TableBodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за минали години)");
            TableBodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            TableBodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            TableBodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            TableBodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            TableBodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            TableBodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квал.");
            TableBodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            TableBodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образованието");
            TableBodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            TableBodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            TableBodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");

            //Брой
            TableBodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc1);
            TableBodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc2);
            TableBodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc3);
            TableBodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc4);
            TableBodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc5);
            TableBodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc6);
            TableBodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc7);
            TableBodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc8);
            TableBodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc9);
            TableBodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc10);
            TableBodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc11);
            TableBodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc12);
            TableBodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc13);
            TableBodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append(Globals.Tab1ReceivedDoc14);
            // Серия
            TableBodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append("C-19");
            TableBodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append("ДС");
            TableBodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append("ОМ-19");
            TableBodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append("ДО");
            TableBodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append("Г-19");
            TableBodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append("ДГ");
            TableBodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append("П-19");
            TableBodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append("ДП");
            TableBodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append("В-19");
            TableBodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append("ДВ");
            TableBodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append("К-19");
            TableBodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append("ДК");
            TableBodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append("СМ-19");
            TableBodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append("О-19");

            //от №
            TableBodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc1Start);
            TableBodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc2Start);
            TableBodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc3Start);
            TableBodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc4Start);
            TableBodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc5Start);
            TableBodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc6Start);
            TableBodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc7Start);
            TableBodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc8Start);
            TableBodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc9Start);
            TableBodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc10Start);
            TableBodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc11Start);
            TableBodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc12Start);
            TableBodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc13Start);
            TableBodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab1Doc14Start);


            //до № 
            TableBodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc1End);
            TableBodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc2End);
            TableBodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc3End);
            TableBodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc4End);
            TableBodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc5End);
            TableBodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc6End);
            TableBodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc7End);
            TableBodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc8End);
            TableBodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc9End);
            TableBodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc10End);
            TableBodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc11End);
            TableBodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc12End);
            TableBodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc13End);
            TableBodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab1Doc14End);

            //Брой
            TableBodyOtchet.Rows[0].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc1);
            TableBodyOtchet.Rows[1].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc2);
            TableBodyOtchet.Rows[2].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc3);
            TableBodyOtchet.Rows[3].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc4);
            TableBodyOtchet.Rows[4].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc5);
            TableBodyOtchet.Rows[5].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc6);
            TableBodyOtchet.Rows[6].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc7);
            TableBodyOtchet.Rows[7].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc8);
            TableBodyOtchet.Rows[8].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc9);
            TableBodyOtchet.Rows[9].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc10);
            TableBodyOtchet.Rows[10].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc11);
            TableBodyOtchet.Rows[11].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc12);
            TableBodyOtchet.Rows[12].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc13);
            TableBodyOtchet.Rows[13].Cells[6].Paragraphs.First().Append(Globals.Tab1ReceivedDoc14);

            //Издадени по рег. книга
            TableBodyOtchet.Rows[0].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc1);
            TableBodyOtchet.Rows[1].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc2);
            TableBodyOtchet.Rows[2].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc3);
            TableBodyOtchet.Rows[3].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc4);
            TableBodyOtchet.Rows[4].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc5);
            TableBodyOtchet.Rows[5].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc6);
            TableBodyOtchet.Rows[6].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc7);
            TableBodyOtchet.Rows[7].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc8);
            TableBodyOtchet.Rows[8].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc9);
            TableBodyOtchet.Rows[9].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc10);
            TableBodyOtchet.Rows[10].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc11);
            TableBodyOtchet.Rows[11].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc12);
            TableBodyOtchet.Rows[12].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc13);
            TableBodyOtchet.Rows[13].Cells[7].Paragraphs.First().Append(Globals.Tab1RegBookDoc14);

            //Анулирани
            TableBodyOtchet.Rows[0].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc1);
            TableBodyOtchet.Rows[1].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc2);
            TableBodyOtchet.Rows[2].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc3);
            TableBodyOtchet.Rows[3].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc4);
            TableBodyOtchet.Rows[4].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc5);
            TableBodyOtchet.Rows[5].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc6);
            TableBodyOtchet.Rows[6].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc7);
            TableBodyOtchet.Rows[7].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc8);
            TableBodyOtchet.Rows[8].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc9);
            TableBodyOtchet.Rows[9].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc10);
            TableBodyOtchet.Rows[10].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc11);
            TableBodyOtchet.Rows[11].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc12);
            TableBodyOtchet.Rows[12].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc13);
            TableBodyOtchet.Rows[13].Cells[8].Paragraphs.First().Append(Globals.Tab1CanceledDoc14);


            //Годни
            TableBodyOtchet.Rows[0].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc1);
            TableBodyOtchet.Rows[1].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc2);
            TableBodyOtchet.Rows[2].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc3);
            TableBodyOtchet.Rows[3].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc4);
            TableBodyOtchet.Rows[4].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc5);
            TableBodyOtchet.Rows[5].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc6);
            TableBodyOtchet.Rows[6].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc7);
            TableBodyOtchet.Rows[7].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc8);
            TableBodyOtchet.Rows[8].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc9);
            TableBodyOtchet.Rows[9].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc10);
            TableBodyOtchet.Rows[10].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc11);
            TableBodyOtchet.Rows[11].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc12);
            TableBodyOtchet.Rows[12].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc13);
            TableBodyOtchet.Rows[13].Cells[9].Paragraphs.First().Append(Globals.Tab1ForDesDoc14);

            //Общ брой унищожени
            TableBodyOtchet.Rows[0].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc1Destroyed);
            TableBodyOtchet.Rows[1].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc2Destroyed);
            TableBodyOtchet.Rows[2].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc3Destroyed);
            TableBodyOtchet.Rows[3].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc4Destroyed);
            TableBodyOtchet.Rows[4].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc5Destroyed);
            TableBodyOtchet.Rows[5].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc6Destroyed);
            TableBodyOtchet.Rows[6].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc7Destroyed);
            TableBodyOtchet.Rows[7].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc8Destroyed);
            TableBodyOtchet.Rows[8].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc9Destroyed);
            TableBodyOtchet.Rows[9].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc10Destroyed);
            TableBodyOtchet.Rows[10].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc11Destroyed);
            TableBodyOtchet.Rows[11].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc12Destroyed);
            TableBodyOtchet.Rows[12].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc13Destroyed);
            TableBodyOtchet.Rows[13].Cells[10].Paragraphs.First().Append(Globals.Tab1Doc14Destroyed);

            //Остатък дубликати
            TableBodyOtchet.Rows[0].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc1Rest);
            TableBodyOtchet.Rows[1].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc2Rest);
            TableBodyOtchet.Rows[2].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc3Rest);
            TableBodyOtchet.Rows[3].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc4Rest);
            TableBodyOtchet.Rows[4].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc5Rest);
            TableBodyOtchet.Rows[5].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc6Rest);
            TableBodyOtchet.Rows[6].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc7Rest);
            TableBodyOtchet.Rows[7].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc8Rest);
            TableBodyOtchet.Rows[8].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc9Rest);
            TableBodyOtchet.Rows[9].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc10Rest);
            TableBodyOtchet.Rows[10].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc11Rest);
            TableBodyOtchet.Rows[11].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc12Rest);
            TableBodyOtchet.Rows[12].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc13Rest);
            TableBodyOtchet.Rows[13].Cells[11].Paragraphs.First().Append(Globals.Tab1Doc14Rest);

            Doc.InsertTable(TableBodyOtchet);
            Doc.Save();
             /*Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
             wordDoc = app.Documents.Open(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
             wordDoc.ExportAsFixedFormat(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
             wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
             app.Quit();
             Marshal.ReleaseComObject(wordDoc);
             Marshal.ReleaseComObject(app);*/
             File.Delete(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
            Process.Start("WINWORD.EXE", fileName1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string fileName1 = "OtchetKod2.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;
            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitleOtchet2 = Doc.AddTable(1, 11);
            TableTitleOtchet2.Alignment = Alignment.center;
            TableTitleOtchet2.Design = TableDesign.TableGrid;
            TableTitleOtchet2.AutoFit = AutoFit.Contents;
            TableTitleOtchet2.Rows[0].MergeCells(2, 5);
            TableTitleOtchet2.Rows[0].MergeCells(4, 5);
            TableTitleOtchet2.Rows[0].Cells[2].Paragraphs.First().Append("Получени от други институции", FormatTitle);
            TableTitleOtchet2.Rows[0].Cells[4].Paragraphs.First().Append("За унищожаване", FormatTitle);


            Doc.InsertTable(TableTitleOtchet2);

            Table TableTitle2Otchet = Doc.AddTable(1, 11);
            TableTitle2Otchet.Alignment = Alignment.center;
            TableTitle2Otchet.Design = TableDesign.TableGrid;
            TableTitle2Otchet.AutoFit = AutoFit.Contents;
            TableTitle2Otchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[2].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[3].Paragraphs.First().Append("от №", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[4].Paragraphs.First().Append("до №", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[5].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[6].Paragraphs.First().Append("Издадени по рег. книга", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[7].Paragraphs.First().Append("Анулирани", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[8].Paragraphs.First().Append("Годни", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[9].Paragraphs.First().Append("Общ брой унищожени", FormatTitle);
            TableTitle2Otchet.Rows[0].Cells[10].Paragraphs.First().Append("Остатък дубликати", FormatTitle);



            Doc.InsertTable(TableTitle2Otchet);

            //Ном. номер
            Table Table2BodyOtchet = Doc.AddTable(14, 11);
            Table2BodyOtchet.Alignment = Alignment.center;
            Table2BodyOtchet.Design = TableDesign.TableGrid;
            Table2BodyOtchet.AutoFit = AutoFit.Fixed;
            Table2BodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            Table2BodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            Table2BodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            Table2BodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            Table2BodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            Table2BodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            Table2BodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            Table2BodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            Table2BodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            Table2BodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            Table2BodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            Table2BodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            Table2BodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            Table2BodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            Table2BodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            Table2BodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            Table2BodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за минали години)");
            Table2BodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table2BodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table2BodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            Table2BodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            Table2BodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            Table2BodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квалификация");
            Table2BodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table2BodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален етап");
            Table2BodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table2BodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            Table2BodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");


            // Серия
            Table2BodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append("C-19");
            Table2BodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append("ДС");
            Table2BodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append("ОМ-19");
            Table2BodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append("ДО");
            Table2BodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append("Г-19");
            Table2BodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append("ДГ");
            Table2BodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append("П-19");
            Table2BodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append("ДП");
            Table2BodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append("В-19");
            Table2BodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append("ДВ");
            Table2BodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append("К-19");
            Table2BodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append("ДК");
            Table2BodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append("СМ-19");
            Table2BodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append("О-19");


            //от №
            Table2BodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc1Start);
            Table2BodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc2Start);
            Table2BodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc3Start);
            Table2BodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc4Start);
            Table2BodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc5Start);
            Table2BodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc6Start);
            Table2BodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc7Start);
            Table2BodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc8Start);
            Table2BodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc9Start);
            Table2BodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc10Start);
            Table2BodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc11Start);
            Table2BodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc12Start);
            Table2BodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc13Start);
            Table2BodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append(Globals.Tab2Doc14Start);

            //до №
            Table2BodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc1End);
            Table2BodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc2End);
            Table2BodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc3End);
            Table2BodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc4End);
            Table2BodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc5End);
            Table2BodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc6End);
            Table2BodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc7End);
            Table2BodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc8End);
            Table2BodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc9End);
            Table2BodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc10End);
            Table2BodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc11End);
            Table2BodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc12End);
            Table2BodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc13End);
            Table2BodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab2Doc14End);


            //Брой
            Table2BodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc1);
            Table2BodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc2);
            Table2BodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc3);
            Table2BodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc4);
            Table2BodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc5);
            Table2BodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc6);
            Table2BodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc7);
            Table2BodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc8);
            Table2BodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc9);
            Table2BodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc10);
            Table2BodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc11);
            Table2BodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc12);
            Table2BodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc13);
            Table2BodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab2ReceivedDoc14);


            //Издадени по рег. книга
            Table2BodyOtchet.Rows[0].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc1);
            Table2BodyOtchet.Rows[1].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc2);
            Table2BodyOtchet.Rows[2].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc3);
            Table2BodyOtchet.Rows[3].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc4);
            Table2BodyOtchet.Rows[4].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc5);
            Table2BodyOtchet.Rows[5].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc6);
            Table2BodyOtchet.Rows[6].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc7);
            Table2BodyOtchet.Rows[7].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc8);
            Table2BodyOtchet.Rows[8].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc9);
            Table2BodyOtchet.Rows[9].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc10);
            Table2BodyOtchet.Rows[10].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc11);
            Table2BodyOtchet.Rows[11].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc12);
            Table2BodyOtchet.Rows[12].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc13);
            Table2BodyOtchet.Rows[13].Cells[6].Paragraphs.First().Append(Globals.Tab2RegBookDoc14);


            //Анулирани
            Table2BodyOtchet.Rows[0].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc1);
            Table2BodyOtchet.Rows[1].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc2);
            Table2BodyOtchet.Rows[2].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc3);
            Table2BodyOtchet.Rows[3].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc4);
            Table2BodyOtchet.Rows[4].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc5);
            Table2BodyOtchet.Rows[5].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc6);
            Table2BodyOtchet.Rows[6].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc7);
            Table2BodyOtchet.Rows[7].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc8);
            Table2BodyOtchet.Rows[8].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc9);
            Table2BodyOtchet.Rows[9].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc10);
            Table2BodyOtchet.Rows[10].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc11);
            Table2BodyOtchet.Rows[11].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc12);
            Table2BodyOtchet.Rows[12].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc13);
            Table2BodyOtchet.Rows[13].Cells[7].Paragraphs.First().Append(Globals.Tab2CanceledDoc14);


            //Годни
            Table2BodyOtchet.Rows[0].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc1);
            Table2BodyOtchet.Rows[1].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc2);
            Table2BodyOtchet.Rows[2].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc3);
            Table2BodyOtchet.Rows[3].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc4);
            Table2BodyOtchet.Rows[4].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc5);
            Table2BodyOtchet.Rows[5].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc6);
            Table2BodyOtchet.Rows[6].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc7);
            Table2BodyOtchet.Rows[7].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc8);
            Table2BodyOtchet.Rows[8].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc9);
            Table2BodyOtchet.Rows[9].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc10);
            Table2BodyOtchet.Rows[10].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc11);
            Table2BodyOtchet.Rows[11].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc12);
            Table2BodyOtchet.Rows[12].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc13);
            Table2BodyOtchet.Rows[13].Cells[8].Paragraphs.First().Append(Globals.Tab2ForDesDoc14);

            //Общ брой унищожени
            Table2BodyOtchet.Rows[0].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc1Destroyed);
            Table2BodyOtchet.Rows[1].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc2Destroyed);
            Table2BodyOtchet.Rows[2].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc3Destroyed);
            Table2BodyOtchet.Rows[3].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc4Destroyed);
            Table2BodyOtchet.Rows[4].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc5Destroyed);
            Table2BodyOtchet.Rows[5].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc6Destroyed);
            Table2BodyOtchet.Rows[6].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc7Destroyed);
            Table2BodyOtchet.Rows[7].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc8Destroyed);
            Table2BodyOtchet.Rows[8].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc9Destroyed);
            Table2BodyOtchet.Rows[9].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc10Destroyed);
            Table2BodyOtchet.Rows[10].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc11Destroyed);
            Table2BodyOtchet.Rows[11].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc12Destroyed);
            Table2BodyOtchet.Rows[12].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc13Destroyed);
            Table2BodyOtchet.Rows[13].Cells[9].Paragraphs.First().Append(Globals.Tab2Doc14Destroyed);


            //Остатък дубликати
            Table2BodyOtchet.Rows[0].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc1Rest);
            Table2BodyOtchet.Rows[1].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc2Rest);
            Table2BodyOtchet.Rows[2].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc3Rest);
            Table2BodyOtchet.Rows[3].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc4Rest);
            Table2BodyOtchet.Rows[4].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc5Rest);
            Table2BodyOtchet.Rows[5].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc6Rest);
            Table2BodyOtchet.Rows[6].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc7Rest);
            Table2BodyOtchet.Rows[7].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc8Rest);
            Table2BodyOtchet.Rows[8].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc9Rest);
            Table2BodyOtchet.Rows[9].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc10Rest);
            Table2BodyOtchet.Rows[10].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc11Rest);
            Table2BodyOtchet.Rows[11].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc12Rest);
            Table2BodyOtchet.Rows[12].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc13Rest);
            Table2BodyOtchet.Rows[13].Cells[10].Paragraphs.First().Append(Globals.Tab2Doc14Rest);



            Doc.InsertTable(Table2BodyOtchet);
            Doc.Save();
            Process.Start("WINWORD.EXE", fileName1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string fileName1 = "OtchetKod3.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;
            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");

            Table TableTitleOtchet3 = Doc.AddTable(1, 6);
            TableTitleOtchet3.Alignment = Alignment.center;
            TableTitleOtchet3.Design = TableDesign.TableGrid;
            TableTitleOtchet3.AutoFit = AutoFit.Contents;
            TableTitleOtchet3.Rows[0].MergeCells(2, 6);
            TableTitleOtchet3.Rows[0].Cells[2].Paragraphs.First().Append("Предадени на други институции", FormatTitle);


            Doc.InsertTable(TableTitleOtchet3);

            Table Table3TitleOtchet = Doc.AddTable(1, 6);
            Table3TitleOtchet.Alignment = Alignment.center;
            Table3TitleOtchet.Design = TableDesign.TableGrid;
            Table3TitleOtchet.AutoFit = AutoFit.Contents;
            Table3TitleOtchet.Rows[0].Cells[0].Paragraphs.First().Append("Ном. номер", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[2].Paragraphs.First().Append("Серия", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[3].Paragraphs.First().Append("от №", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[4].Paragraphs.First().Append("до №", FormatTitle);
            Table3TitleOtchet.Rows[0].Cells[5].Paragraphs.First().Append("Брой", FormatTitle);



            Doc.InsertTable(Table3TitleOtchet);


            Table Table3BodyOtchet = Doc.AddTable(14, 6);
            Table3BodyOtchet.Alignment = Alignment.center;
            Table3BodyOtchet.Design = TableDesign.TableGrid;
            Table3BodyOtchet.AutoFit = AutoFit.Fixed;
            Table3BodyOtchet.Rows[0].Cells[0].Paragraphs.First().Append("3-34", FormatTitle);
            Table3BodyOtchet.Rows[1].Cells[0].Paragraphs.First().Append("3-44a", FormatTitle);
            Table3BodyOtchet.Rows[2].Cells[0].Paragraphs.First().Append("3-20", FormatTitle);
            Table3BodyOtchet.Rows[3].Cells[0].Paragraphs.First().Append("3-30a", FormatTitle);
            Table3BodyOtchet.Rows[4].Cells[0].Paragraphs.First().Append("3-22", FormatTitle);
            Table3BodyOtchet.Rows[5].Cells[0].Paragraphs.First().Append("3-22a", FormatTitle);
            Table3BodyOtchet.Rows[6].Cells[0].Paragraphs.First().Append("3-54", FormatTitle);
            Table3BodyOtchet.Rows[7].Cells[0].Paragraphs.First().Append("3-54а", FormatTitle);
            Table3BodyOtchet.Rows[8].Cells[0].Paragraphs.First().Append("3-54B", FormatTitle);
            Table3BodyOtchet.Rows[9].Cells[0].Paragraphs.First().Append("3-54aB", FormatTitle);
            Table3BodyOtchet.Rows[10].Cells[0].Paragraphs.First().Append("3-27B", FormatTitle);
            Table3BodyOtchet.Rows[11].Cells[0].Paragraphs.First().Append("3-27aB", FormatTitle);
            Table3BodyOtchet.Rows[12].Cells[0].Paragraphs.First().Append("3-42", FormatTitle);
            Table3BodyOtchet.Rows[13].Cells[0].Paragraphs.First().Append("3-30", FormatTitle);
            //Наименование на документа
            Table3BodyOtchet.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            Table3BodyOtchet.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на  диплома за средно образование");
            Table3BodyOtchet.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование(за мин. год.)");
            Table3BodyOtchet.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table3BodyOtchet.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table3BodyOtchet.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен  гимназиален етап");
            Table3BodyOtchet.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за професионална квалификация");
            Table3BodyOtchet.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за професионална квалификация");
            Table3BodyOtchet.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за  валидиране професионална квалификация");
            Table3BodyOtchet.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table3BodyOtchet.Rows[10].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образованието");
            Table3BodyOtchet.Rows[11].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table3BodyOtchet.Rows[12].Cells[1].Paragraphs.First().Append("Диплома за средно образование образец за минали години");
            Table3BodyOtchet.Rows[13].Cells[1].Paragraphs.First().Append("Свидетелство за основно образование");
            // Серия
            Table3BodyOtchet.Rows[0].Cells[2].Paragraphs.First().Append("C-19");
            Table3BodyOtchet.Rows[1].Cells[2].Paragraphs.First().Append("ДС");
            Table3BodyOtchet.Rows[2].Cells[2].Paragraphs.First().Append("ОМ-19");
            Table3BodyOtchet.Rows[3].Cells[2].Paragraphs.First().Append("ДО");
            Table3BodyOtchet.Rows[4].Cells[2].Paragraphs.First().Append("Г-19");
            Table3BodyOtchet.Rows[5].Cells[2].Paragraphs.First().Append("ДГ");
            Table3BodyOtchet.Rows[6].Cells[2].Paragraphs.First().Append("П-19");
            Table3BodyOtchet.Rows[7].Cells[2].Paragraphs.First().Append("ДП");
            Table3BodyOtchet.Rows[8].Cells[2].Paragraphs.First().Append("В-19");
            Table3BodyOtchet.Rows[9].Cells[2].Paragraphs.First().Append("ДВ");
            Table3BodyOtchet.Rows[10].Cells[2].Paragraphs.First().Append("К-19");
            Table3BodyOtchet.Rows[11].Cells[2].Paragraphs.First().Append("ДК");
            Table3BodyOtchet.Rows[12].Cells[2].Paragraphs.First().Append("СМ-19");
            Table3BodyOtchet.Rows[13].Cells[2].Paragraphs.First().Append("О-19");

            //от №
            Table3BodyOtchet.Rows[0].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc1Start);
            Table3BodyOtchet.Rows[1].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc2Start);
            Table3BodyOtchet.Rows[2].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc3Start);
            Table3BodyOtchet.Rows[3].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc4Start);
            Table3BodyOtchet.Rows[4].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc5Start);
            Table3BodyOtchet.Rows[5].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc6Start);
            Table3BodyOtchet.Rows[6].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc7Start);
            Table3BodyOtchet.Rows[7].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc8Start);
            Table3BodyOtchet.Rows[8].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc9Start);
            Table3BodyOtchet.Rows[9].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc10Start);
            Table3BodyOtchet.Rows[10].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc11Start);
            Table3BodyOtchet.Rows[11].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc12Start);
            Table3BodyOtchet.Rows[12].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc13Start);
            Table3BodyOtchet.Rows[13].Cells[3].Paragraphs.First().Append(Globals.Tab3Doc14Start);


            //до № 
            Table3BodyOtchet.Rows[0].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc1End);
            Table3BodyOtchet.Rows[1].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc2End);
            Table3BodyOtchet.Rows[2].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc3End);
            Table3BodyOtchet.Rows[3].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc4End);
            Table3BodyOtchet.Rows[4].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc5End);
            Table3BodyOtchet.Rows[5].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc6End);
            Table3BodyOtchet.Rows[6].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc7End);
            Table3BodyOtchet.Rows[7].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc8End);
            Table3BodyOtchet.Rows[8].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc9End);
            Table3BodyOtchet.Rows[9].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc10End);
            Table3BodyOtchet.Rows[10].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc11End);
            Table3BodyOtchet.Rows[11].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc12End);
            Table3BodyOtchet.Rows[12].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc13End);
            Table3BodyOtchet.Rows[13].Cells[4].Paragraphs.First().Append(Globals.Tab3Doc14End);


            //Брой
            Table3BodyOtchet.Rows[0].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc1Count);
            Table3BodyOtchet.Rows[1].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc2Count);
            Table3BodyOtchet.Rows[2].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc3Count);
            Table3BodyOtchet.Rows[3].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc4Count);
            Table3BodyOtchet.Rows[4].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc5Count);
            Table3BodyOtchet.Rows[5].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc6Count);
            Table3BodyOtchet.Rows[6].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc7Count);
            Table3BodyOtchet.Rows[7].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc8Count);
            Table3BodyOtchet.Rows[8].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc9Count);
            Table3BodyOtchet.Rows[9].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc10Count);
            Table3BodyOtchet.Rows[10].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc11Count);
            Table3BodyOtchet.Rows[11].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc12Count);
            Table3BodyOtchet.Rows[12].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc13Count);
            Table3BodyOtchet.Rows[13].Cells[5].Paragraphs.First().Append(Globals.Tab3Doc14Count);





            Doc.InsertTable(Table3BodyOtchet);
            Doc.Save();
            Process.Start("WINWORD.EXE", fileName1);
        }

        private void button7_Click(object sender, EventArgs e)
        {

            string fileName1 = "OtchetKod2.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;
            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess1 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess2 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess3 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess4 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess5 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess6 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess7 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess8 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess9 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess10 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess11 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess12 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess13 = new Database_Elements.DataAccess();
            int count334a_2018 = dataAccess.GetCountDupsForYear("3-34a","2018");
            int count330a_2018 = dataAccess1.GetCountDupsForYear("3-30a", "2018");
            int count32a_2018 = dataAccess2.GetCountDupsForYear("3-22a", "2018");
            int count354a_2018 = dataAccess3.GetCountDupsForYear("3-54a", "2018");
            int count354aB_2018 = dataAccess4.GetCountDupsForYear("3-54aB", "2018");
            int count327aB_2018 = dataAccess5.GetCountDupsForYear("3-27aB", "2018");
            int count344a_2018 = dataAccess6.GetCountDupsForYear("3-44a", "2018");
            int count334a_2019 = dataAccess7.GetCountDupsForYear("3-34a", "2019");
            int count330a_2019 = dataAccess8.GetCountDupsForYear("3-30a", "2019");
            int count32a_2019 = dataAccess9.GetCountDupsForYear("3-22a", "2019");
            int count354a_2019 = dataAccess10.GetCountDupsForYear("3-54a", "2019");
            int count354aB_2019 = dataAccess11.GetCountDupsForYear("3-54aB", "2019");
            int count327aB_2019 = dataAccess12.GetCountDupsForYear("3-27aB", "2019");
            int count344a_2019 = dataAccess13.GetCountDupsForYear("3-44a", "2019");

            Table TableTitle4 = Doc.AddTable(1, 6);
            TableTitle4.Design = TableDesign.TableGrid;
            TableTitle4.AutoFit = AutoFit.Contents;
            TableTitle4.Rows[0].Cells[0].Paragraphs.First().Append("Година на получаване", FormatTitle);
            TableTitle4.Rows[0].Cells[1].Paragraphs.First().Append("Номенклатурен №", FormatTitle);
            TableTitle4.Rows[0].Cells[2].Paragraphs.First().Append("Наименование на дубликата", FormatTitle);
            TableTitle4.Rows[0].Cells[3].Paragraphs.First().Append("Брой", FormatTitle);
            TableTitle4.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
            TableTitle4.Rows[0].Cells[5].Paragraphs.First().Append("Описание на фабричните номера на наличните дубликати", FormatTitle);

            Doc.InsertTable(TableTitle4);
            Table Table4 = Doc.AddTable(8, 6);
            //Година на получаване
            Table4.Rows[0].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[1].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[2].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[3].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[4].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[5].Cells[0].Paragraphs.First().Append("2018");
            Table4.Rows[6].Cells[0].Paragraphs.First().Append("2019");
            Table4.Rows[7].Cells[0].Paragraphs.First().Append("2019");


            //ном.номер
            Table4.Rows[0].Cells[1].Paragraphs.First().Append("3-34а");
            Table4.Rows[1].Cells[1].Paragraphs.First().Append("3-30а");
            Table4.Rows[2].Cells[1].Paragraphs.First().Append("3-22а");
            Table4.Rows[3].Cells[1].Paragraphs.First().Append("3-54а");
            Table4.Rows[4].Cells[1].Paragraphs.First().Append("3-54аВ");
            Table4.Rows[5].Cells[1].Paragraphs.First().Append("3-27аВ");
            Table4.Rows[6].Cells[1].Paragraphs.First().Append("3-44а");
            Table4.Rows[7].Cells[1].Paragraphs.First().Append("3-30а");


            //наименование на дубл.
            Table4.Rows[0].Cells[2].Paragraphs.First().Append("Дубликат на диплома");
            Table4.Rows[1].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table4.Rows[2].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            Table4.Rows[3].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за проф. квалификация");
            Table4.Rows[4].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table4.Rows[5].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за валид. на комп.");
            Table4.Rows[6].Cells[2].Paragraphs.First().Append("Дубликат на диплома");
            Table4.Rows[7].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");


            //брой
            Table4.Rows[0].Cells[3].Paragraphs.First().Append(count334a_2018.ToString());
            Table4.Rows[1].Cells[3].Paragraphs.First().Append(count330a_2018.ToString());
            Table4.Rows[2].Cells[3].Paragraphs.First().Append(count32a_2018.ToString());
            Table4.Rows[3].Cells[3].Paragraphs.First().Append(count354a_2018.ToString());
            Table4.Rows[4].Cells[3].Paragraphs.First().Append(count354aB_2018.ToString());
            Table4.Rows[5].Cells[3].Paragraphs.First().Append(count327aB_2018.ToString());
            Table4.Rows[6].Cells[3].Paragraphs.First().Append(count344a_2019.ToString());
            Table4.Rows[7].Cells[3].Paragraphs.First().Append(count330a_2019.ToString());


            //серия
            Table4.Rows[0].Cells[4].Paragraphs.First().Append("ДС");
            Table4.Rows[1].Cells[4].Paragraphs.First().Append("ДО");
            Table4.Rows[2].Cells[4].Paragraphs.First().Append("ДГ");
            Table4.Rows[3].Cells[4].Paragraphs.First().Append("ДП");
            Table4.Rows[4].Cells[4].Paragraphs.First().Append("ДВ");
            Table4.Rows[5].Cells[4].Paragraphs.First().Append("ДК");
            Table4.Rows[6].Cells[4].Paragraphs.First().Append("ДС");
            Table4.Rows[7].Cells[4].Paragraphs.First().Append("ДО");


            //остатък
            Table4.Rows[0].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[1].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[2].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[3].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[4].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[5].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[6].Cells[5].Paragraphs.First().Append("");
            Table4.Rows[7].Cells[5].Paragraphs.First().Append("");




            Doc.InsertTable(Table4);

            Table Table4a = Doc.AddTable(4, 6);
            //година
            Table4a.Rows[0].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[1].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[2].Cells[0].Paragraphs.First().Append("2019");
            Table4a.Rows[3].Cells[0].Paragraphs.First().Append("2019");
            //ном.номер
            Table4a.Rows[0].Cells[1].Paragraphs.First().Append("3-54а");
            Table4a.Rows[1].Cells[1].Paragraphs.First().Append("3-54аВ");
            Table4a.Rows[2].Cells[1].Paragraphs.First().Append("3-27аВ");
            Table4a.Rows[3].Cells[1].Paragraphs.First().Append("3-22а");
            //наименование
            Table4a.Rows[0].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за проф. квалификация");
            Table4a.Rows[1].Cells[2].Paragraphs.First().Append("Дубликат на свидетелство за валидиране");
            Table4a.Rows[2].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table4a.Rows[3].Cells[2].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            //брой
            Table4a.Rows[0].Cells[3].Paragraphs.First().Append(count32a_2019.ToString());
            Table4a.Rows[1].Cells[3].Paragraphs.First().Append(count354a_2019.ToString());
            Table4a.Rows[2].Cells[3].Paragraphs.First().Append(count354aB_2019.ToString());
            Table4a.Rows[3].Cells[3].Paragraphs.First().Append(count327aB_2019.ToString());
            //серия
            Table4a.Rows[0].Cells[4].Paragraphs.First().Append("ДП");
            Table4a.Rows[1].Cells[4].Paragraphs.First().Append("ДВ");
            Table4a.Rows[2].Cells[4].Paragraphs.First().Append("ДК");
            Table4a.Rows[3].Cells[4].Paragraphs.First().Append("ДГ");
            //остатък
            Table4a.Rows[0].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[1].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[2].Cells[5].Paragraphs.First().Append("");
            Table4a.Rows[3].Cells[5].Paragraphs.First().Append("");

            Doc.InsertTable(Table4a);

            Doc.Save();
            Process.Start("WINWORD.EXE", fileName1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName1 = "Table5.docx";
            var Doc = DocX.Create(fileName1);
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Portrait;
            Database_Elements.DataAccess dataAccess = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess1 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess2 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess3 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess4 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess5 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess6 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess7 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess8 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess9 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess10 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess11 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess12 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess13 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess14 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess15 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess16 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess17 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess18 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess19 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess20 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess21 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess22 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess23 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess24 = new Database_Elements.DataAccess();
            Database_Elements.DataAccess dataAccess25 = new Database_Elements.DataAccess();
            int count334 = dataAccess.GetCountOfDestroyedDoc("3-34");
            int count344a = dataAccess1.GetCountOfDestroyedDoc("3-44a");
            int count320 = dataAccess2.GetCountOfDestroyedDoc("3-20");
            int count330a = dataAccess3.GetCountOfDestroyedDoc("3-30a");
            int count322 = dataAccess4.GetCountOfDestroyedDoc("3-22");
            int count322a = dataAccess5.GetCountOfDestroyedDoc("3-22a");
            int count354 = dataAccess6.GetCountOfDestroyedDoc("3-54");
            int count354a = dataAccess7.GetCountOfDestroyedDoc("3-54a");
            int count354B = dataAccess8.GetCountOfDestroyedDoc("3-54B");
            int count354aB = dataAccess9.GetCountOfDestroyedDoc("3-54aB");
            int count327B = dataAccess10.GetCountOfDestroyedDoc("3-27B");
            int count327aB = dataAccess11.GetCountOfDestroyedDoc("3-27aB");
            int count342 = dataAccess12.GetCountOfDestroyedDoc("3-42");
            int count334a = dataAccess13.GetCountOfDestroyedDoc("3-34a");
            Formatting FormatTitle = new Formatting();
            FormatTitle.Size = 11D;
            FormatTitle.Bold = true;
            FormatTitle.Position = 20;
            FormatTitle.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph = Doc.InsertParagraph("ПРЕДАВА СЕ В РУО ЗАВЕРЕНО КОПИЕ", false, FormatTitle);
            Titleparagraph.Alignment = Alignment.right;

            Paragraph empty1 = Doc.InsertParagraph("");
            string schoolName = dataAccess14.GetSchoolName();
            Paragraph Titleparagraph1 = Doc.InsertParagraph(schoolName, false, FormatTitle);
            Titleparagraph1.Alignment = Alignment.right;
            string schoolNeispuo = dataAccess15.GetSchoolNeispuo();
            string street = dataAccess16.GetSchoolStreet();
            string city = dataAccess17.GetSchoolCity();
            string community = dataAccess18.GetSchoolCommunity();
            string area = dataAccess19.GetSchoolArea();
            string postCode = dataAccess20.GetSchoolPostCode();
            string phone = dataAccess21.GetSchoolPhone();
            string email = dataAccess22.GetSchoolEmail();
            string principalName = dataAccess23.GetSchoolPrincipalName();
            string manager = dataAccess24.GetManager();

            Paragraph Titleparagraph2 = Doc.InsertParagraph($"*{postCode} {street},община {community},област {area}");
            Titleparagraph2.Alignment = Alignment.center;


            Paragraph Titleparagraph3 = Doc.InsertParagraph($"гр.(с.){city}, {phone}, {email}, {principalName}");
            Titleparagraph3.Alignment = Alignment.center;


            Paragraph Titleparagraph4 = Doc.InsertParagraph("П Р О Т О К О Л", false, FormatTitle);
            Titleparagraph4.Alignment = Alignment.center;

            Paragraph empty2 = Doc.InsertParagraph("");

            Paragraph Titleparagraph5 = Doc.InsertParagraph($"Днес, 20.04.2020, комисия, назначена със Заповед № {Globals.OrderNumber} от {Globals.OrderDate.ToString()} на директора в състав:");
            Titleparagraph5.Alignment = Alignment.left;

            Paragraph Titleparagraph6 = Doc.InsertParagraph($"Председател:{manager}");
            Titleparagraph6.Alignment = Alignment.left;

            Paragraph Titleparagraph7 = Doc.InsertParagraph("Членове:");
            Titleparagraph7.Alignment = Alignment.left;

            SqlConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand("SELECT * FROM department_members WHERE id=id", connection);
                SqlDataReader reader = command.ExecuteReader();
                int i = 1;
                using (reader)
                {
                    while (reader.Read())
                    {
                        string fullName = (string)reader["first_name"] + " " + (string)reader["middle_name"] + " " + (string)reader["last_name"];
                        Doc.InsertParagraph($"{i}. {fullName}");
                    }
                }
            }
            connection.Close();

            Paragraph empty3 = Doc.InsertParagraph("");
            Doc.InsertParagraph("Комисия извърши унищожаването на следните използвани, сгрешени при попълването им,както и остатъкът от неизползвани документи с фабрична номерация, съответно за 2018/2019 и дубликати от предходни учеби годинр както следва:");
            Doc.InsertParagraph("");
            Table TableTitle5 = Doc.AddTable(1, 3);
            TableTitle5.Alignment = Alignment.left;
            TableTitle5.Design = TableDesign.TableGrid;
            TableTitle5.AutoFit = AutoFit.Contents;
            TableTitle5.Rows[0].Cells[0].Paragraphs.First().Append("Номенклатурен №", FormatTitle);
            TableTitle5.Rows[0].Cells[1].Paragraphs.First().Append("Наименование на документа", FormatTitle);
            TableTitle5.Rows[0].Cells[2].Paragraphs.First().Append("Общ брой унищожени", FormatTitle);


            Doc.InsertTable(TableTitle5);
            Table Table5 = Doc.AddTable(10, 3);
            Table5.Alignment = Alignment.left;
            Table5.Rows[0].Cells[0].Paragraphs.First().Append("3-34");
            Table5.Rows[1].Cells[0].Paragraphs.First().Append("3-44а");
            Table5.Rows[2].Cells[0].Paragraphs.First().Append("3-20");
            Table5.Rows[3].Cells[0].Paragraphs.First().Append("3-30а");
            Table5.Rows[4].Cells[0].Paragraphs.First().Append("3-22");
            Table5.Rows[5].Cells[0].Paragraphs.First().Append("3-22а");
            Table5.Rows[6].Cells[0].Paragraphs.First().Append("3-54");
            Table5.Rows[7].Cells[0].Paragraphs.First().Append("3-54a");
            Table5.Rows[8].Cells[0].Paragraphs.First().Append("3-54В");
            Table5.Rows[9].Cells[0].Paragraphs.First().Append("3-54aВ");



            Table5.Rows[0].Cells[1].Paragraphs.First().Append("Диплома за средно образование");
            Table5.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на диплома");
            Table5.Rows[2].Cells[1].Paragraphs.First().Append("Свидетелство за основно");
            Table5.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за основно образование");
            Table5.Rows[4].Cells[1].Paragraphs.First().Append("Удостоверение за завършен гимназиален етап");
            Table5.Rows[5].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за завършен гимназиален етап");
            Table5.Rows[6].Cells[1].Paragraphs.First().Append("Свидетелство за проф. квалификация");
            Table5.Rows[7].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за проф. квалификация");
            Table5.Rows[8].Cells[1].Paragraphs.First().Append("Свидетелство за валидиране на проф. квалификация");
            Table5.Rows[9].Cells[1].Paragraphs.First().Append("Дубликат на свидетелство за валидиране	");




            Table5.Rows[0].Cells[2].Paragraphs.First().Append(count334.ToString()) ;
            Table5.Rows[1].Cells[2].Paragraphs.First().Append(count344a.ToString());
            Table5.Rows[2].Cells[2].Paragraphs.First().Append(count320.ToString());
            Table5.Rows[3].Cells[2].Paragraphs.First().Append(count330a.ToString());
            Table5.Rows[4].Cells[2].Paragraphs.First().Append(count322.ToString());
            Table5.Rows[5].Cells[2].Paragraphs.First().Append(count322a.ToString());
            Table5.Rows[6].Cells[2].Paragraphs.First().Append(count354.ToString());
            Table5.Rows[7].Cells[2].Paragraphs.First().Append(count354a.ToString());
            Table5.Rows[8].Cells[2].Paragraphs.First().Append(count354B.ToString());
            Table5.Rows[9].Cells[2].Paragraphs.First().Append(count354aB.ToString());



            Doc.InsertTable(Table5);

            



            Table Table5end = Doc.AddTable(4, 3);
            Table5end.Alignment = Alignment.left;

            Table5end.Rows[0].Cells[0].Paragraphs.First().Append("3-27В");
            Table5end.Rows[1].Cells[0].Paragraphs.First().Append("3-27aВ");
            Table5end.Rows[2].Cells[0].Paragraphs.First().Append("3-42");
            Table5end.Rows[3].Cells[0].Paragraphs.First().Append("3-34a");

            Table5end.Rows[0].Cells[1].Paragraphs.First().Append("Удостоверение за валидиране на компетентности за начален или първи гимназиален етап/основна степен на образование");
            Table5end.Rows[1].Cells[1].Paragraphs.First().Append("Дубликат на удостоверение за валидиране на компетентности");
            Table5end.Rows[2].Cells[1].Paragraphs.First().Append("Диплома за средно образование - образец за минали");
            Table5end.Rows[3].Cells[1].Paragraphs.First().Append("Дубликат на диплома");

            Table5end.Rows[0].Cells[2].Paragraphs.First().Append(count327B.ToString());
            Table5end.Rows[1].Cells[2].Paragraphs.First().Append(count327aB.ToString());
            Table5end.Rows[2].Cells[2].Paragraphs.First().Append(count342.ToString());
            Table5end.Rows[3].Cells[2].Paragraphs.First().Append(count334a.ToString());

            Doc.InsertTable(Table5end);


            
            Doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;

            Formatting FormatTitle1 = new Formatting();
            FormatTitle1.Size = 11D;
            FormatTitle1.Bold = true;
            FormatTitle1.Position = 20;
            FormatTitle1.FontFamily = new Xceed.Document.NET.Font("Times new roman");
            Paragraph Titleparagraph9 = Doc.InsertParagraph("ПРЕДАВА СЕ В РУО ЗАВЕРЕНО КОПИЕ", false, FormatTitle);
            Titleparagraph.Alignment = Alignment.right;

            Doc.InsertParagraph("");

            if(count334>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година на получаване", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count334, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-34\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-34");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("С-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("C-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count344a>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година на получаване", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count344a, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-44a\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-44a");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДС");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДС " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count320>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година на получаване", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count320, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-20\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-20");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ОМ-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ОМ-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count330a>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count330a, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-30a\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-30a");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДO");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДO " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count322>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count322, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-22\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-22");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("Г-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("Г-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count322a>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименованиe", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count322a, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-22a\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-22a");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДГ");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДГ " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count354>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count354, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-54\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-54");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("П-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("П-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }
            if(count354a>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count354a, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-54a\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-54a");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДП");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДП " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count354B>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count354B, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-54B\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-54B");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("В-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("В-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count354aB>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count354aB, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-54aB\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-54aB");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДВ");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДВ " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count327B>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count327B, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-27B\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-27B");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("К-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("К-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count327aB>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count327aB, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-27aB\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-27aB");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДК");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДК " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }

            if(count342>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count342, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-42\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-42");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("СМ-19");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("СМ-19 " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }
            if(count334a>0)
            {
                Table TableTitle6 = Doc.AddTable(1, 7);
                TableTitle6.Alignment = Alignment.left;
                TableTitle6.Design = TableDesign.TableGrid;
                TableTitle6.AutoFit = AutoFit.Contents;
                TableTitle6.Rows[0].Cells[0].Paragraphs.First().Append("№ по ред", FormatTitle);
                TableTitle6.Rows[0].Cells[1].Paragraphs.First().Append("Година", FormatTitle);
                TableTitle6.Rows[0].Cells[2].Paragraphs.First().Append("Номенклатурен №	", FormatTitle);
                TableTitle6.Rows[0].Cells[3].Paragraphs.First().Append("Наименование", FormatTitle);
                TableTitle6.Rows[0].Cells[4].Paragraphs.First().Append("Серия", FormatTitle);
                TableTitle6.Rows[0].Cells[5].Paragraphs.First().Append("Номер", FormatTitle);
                TableTitle6.Rows[0].Cells[6].Paragraphs.First().Append("О Т Р Я З Ъ К", FormatTitle);
                Doc.InsertTable(TableTitle6);

                Table Table6 = Doc.AddTable(count334a, 7);
                Table6.Alignment = Alignment.left;
                Table6.AutoFit = AutoFit.Contents;

                SqlConnection connection2 = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
                connection2.Open();
                using (connection2)
                {
                    SqlCommand command = new SqlCommand("SELECT * FROM \"3-34a\" WHERE status='Годен за унищожаване' OR status='Анулиран'", connection2);
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    using (reader)
                    {
                        while (reader.Read())
                        {
                            Table6.Rows[i].Cells[0].Paragraphs.First().Append((i + 1).ToString());
                            Table6.Rows[i].Cells[1].Paragraphs.First().Append((string)reader["year"]);
                            Table6.Rows[i].Cells[2].Paragraphs.First().Append("3-34a");
                            Table6.Rows[i].Cells[3].Paragraphs.First().Append((string)reader["name"]);
                            Table6.Rows[i].Cells[4].Paragraphs.First().Append("ДС");
                            Table6.Rows[i].Cells[5].Paragraphs.First().Append((string)reader["fabric_number"]);
                            Table6.Rows[i].Cells[6].Paragraphs.First().Append("ДС " + "№                                                                                              ");
                            i++;
                        }
                    }
                }
                Doc.InsertTable(Table6);
                Doc.InsertParagraph("");
            }
            Paragraph paragraph1 = Doc.InsertParagraph("                               КОМИСИЯ:                                                            Членове:                                       ", false, FormatTitle);
            paragraph1.Alignment = Alignment.left;
            Paragraph paragraph2 = Doc.InsertParagraph("Председател:.......................                                            ........................                Директор:........................                ", false, FormatTitle);
            paragraph2.Alignment = Alignment.center;
            Paragraph paragraph3 = Doc.InsertParagraph("(подпис)                                                                                  (подпис)                                         (подпис и печат)", false, FormatTitle);
            paragraph3.Alignment = Alignment.center;

            Doc.Save();
            Process.Start("WINWORD.EXE", fileName1);
        }
    }
}
