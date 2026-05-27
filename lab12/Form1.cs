using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace lab12
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            try
            {
                string templatePath = @"C:\Users\Антон\Desktop\business_letter.docx";

                Word.Application wordApp = new Word.Application();

                Word.Document doc = wordApp.Documents.Add(templatePath);

                FindAndReplace(wordApp, "name", txtName.Text);
                FindAndReplace(wordApp, "company", txtCompany.Text);
                FindAndReplace(wordApp, "address", txtAddress.Text);
                FindAndReplace(wordApp, "phone", txtPhone.Text);
                FindAndReplace(wordApp, "email", txtEmail.Text);
                FindAndReplace(wordApp, "message", txtMessage.Text);
                FindAndReplace(wordApp, "date", DateTime.Now.ToShortDateString());

                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }

        private void FindAndReplace(Word.Application wordApp, object findText, object replaceText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(
                ref findText,
                ref matchCase,
                ref matchWholeWord,
                ref matchWildCards,
                ref matchSoundsLike,
                ref matchAllWordForms,
                ref forward,
                ref wrap,
                ref format,
                ref replaceText,
                ref replace,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl
            );
        }
    }
}