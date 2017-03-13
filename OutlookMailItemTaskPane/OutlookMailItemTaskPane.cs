using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookMailItemTaskPane
{
    public partial class OutlookMailItemTaskPane : UserControl
    {
        public OutlookMailItemTaskPane()
        {
            InitializeComponent();
            radioButton2.Select();
            radioButton13.Select();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Outlook.MailItem selection = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;

            if (selection == null) return;
            var mailItem = selection;

            if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatPlain)
            {
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
                mailItem.Save();
            }

            var isChecked = checkBox1.Checked ? "YES" : "NO";
            mailItem.To = "BLA.Fileroom@sonymusic.com";

            var
            insert = $"Artist Name: {textBox1.Text + Environment.NewLine}";
            insert += $"Rep Owner/Label: {radioGroup1.Button.Text + Environment.NewLine}";
            insert += $"Contracted With: {textBox2.Text + Environment.NewLine}";
            insert += $"Contract Date: {dateTimePicker1.Text + Environment.NewLine}";
            insert += $"Author: {textBox7.Text + Environment.NewLine}";
            insert += $"Document#: {textBox3.Text + Environment.NewLine}";
            insert += $"Hummingbird#: {textBox4.Text + Environment.NewLine}";
            insert += $"Selection/Project#: {textBox5.Text + Environment.NewLine}";
            insert += $"Subject/Comments: {textBox6.Text + Environment.NewLine}";
            insert += $"Send Copy to GDB Digital File Room: {isChecked + Environment.NewLine}";
            insert += $"Document Type: {radioGroup2.Button.Text + Environment.NewLine}";
            insert += string.Format(Environment.NewLine);

            if (mailItem.BodyFormat != Outlook.OlBodyFormat.olFormatPlain)
            {
                var wd = mailItem.GetInspector.WordEditor as Word.Document;

                if (wd != null)
                {
                    wd.Paragraphs[1].Range.InsertBefore(insert);

                    var rstring = new[] { "Artist Name: ", "Rep Owner/Label: ", "Contracted With: ", 
                        "Contract Date: ", "Author: ", "Document#: ", "Hummingbird#: ", "Selection/Project#: ",
                        "Subject/Comments: ", "Send Copy to GDB Digital File Room: ", "Document Type: " };

                    foreach (var s in rstring)
                    {
                        ReplaceFont(wd.Range(), s);
                    }

                    mailItem.Save();
                }
            }

            Helper.ClearFormControls(this);
            radioButton2.Select();
            radioButton13.Select();
        }

        private static bool ReplaceFont(Word.Range rng, string findWhat)
        {
            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();
            //rng.Find.Replacement.Font.ColorIndex = Word.WdColorIndex.wdBlue;
            rng.Find.Replacement.Font.Bold = -1;
            rng.Find.Text = findWhat;
            rng.Find.Replacement.Text = findWhat;
            rng.Find.Forward = true ;
            rng.Find.Wrap = Word.WdFindWrap.wdFindStop;

            //change this property to true as we want to replace format
            rng.Find.Format = true;

            var hasFound = rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            return hasFound;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }

        private void OutlookMailItemTaskPane_Load(object sender, EventArgs e)
        {
            Focus();
        }

        private void label3_Click(object sender, EventArgs e)
        {
        }
    }

    //internal class Helper
    //{
    //    public static void ClearFormControls(OutlookMailItemTaskPane form)
    //    {
    //        foreach (Control control in form.Controls)
    //        {
    //            var box = control as TextBox;
    //            if (box != null)
    //            {
    //                var txtbox = box;
    //                txtbox.Text = string.Empty;
    //            }
    //            else
    //            {
    //                var checkBox = control as CheckBox;
    //                if (checkBox != null)
    //                {
    //                    var chkbox = checkBox;
    //                    chkbox.Checked = false;
    //                }
    //                else
    //                {
    //                    var button = control as RadioButton;
    //                    if (button != null)
    //                    {
    //                        var rdbtn = button;
    //                        rdbtn.Checked = false;
    //                    }
    //                    else
    //                    {
    //                        var picker = control as DateTimePicker;
    //                        if (picker == null) continue;
    //                        var dtp = picker;
    //                        dtp.Value = DateTime.Now;
    //                    }
    //                }
    //            }
    //        }
    //    }
    //}
}