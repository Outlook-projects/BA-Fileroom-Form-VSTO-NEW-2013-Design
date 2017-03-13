using System;
using System.Windows.Forms;
using OutlookMailItemTaskPane;

namespace OutlookMailItemTaskPane
{
    public partial class TaskPaneControl : UserControl
    {
        public TaskPaneControl()
        {
            InitializeComponent();
        }

        private void TaskPaneControl_Enter(object sender, EventArgs e)
        {
           // textBox1.Focus();
            base.OnEnter(e);
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.MailItem selection = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;

            if (selection == null) return;
            var mailItem = selection;
            var bodyString = mailItem.HTMLBody;
            const string nbs = @"&nbsp;&nbsp;&nbsp;";

            mailItem.HTMLBody = String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Artist Name</b>:{0}{1}</span>", nbs, textBox1.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Rep Owner/Label</b>:{0}{1}</span>", nbs, radioGroup1.Button.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Contracted With</b>:{0}{1}</span>", nbs, textBox2.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Contract Date</b>:{0}{1}</span>", nbs, radDateTimePicker1.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Document#</b>:{0}{1}</span>", nbs, textBox3.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Hummingbird#</b>:{0}{1}</span>", nbs, textBox4.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Section/Project#</b>:{0}{1}</span>", nbs, textBox5.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Subject/Comments</b>:{0}{1}</span>", nbs, textBox6.Text);
            mailItem.HTMLBody += String.Format("<span style='font-size:11.0pt;font-family:Calibri Light; color:#1F4E79;'><b>Document Type</b>:{0}{1}</span>", nbs, radioGroup2.Button.Text);
            mailItem.HTMLBody += String.Format("");

            mailItem.HTMLBody += bodyString;
        }

        private void radButton2_Click(object sender, EventArgs e)
        {
            var aboutBox = new RadAboutBox1();
           // aboutBox.ShowDialog();

        }
    }
}
