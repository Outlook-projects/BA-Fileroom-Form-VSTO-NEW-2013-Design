using System;
using System.Windows.Forms;

namespace OutlookMailItemTaskPane
{
    internal class Helper
    {
        public static void ClearFormControls(OutlookMailItemTaskPane form)
        {
            foreach (Control control in form.Controls)
            {
                var box = control as TextBox;
                if (box != null)
                {
                    var txtbox = box;
                    txtbox.Text = string.Empty;
                }
                else
                {
                    var checkBox = control as CheckBox;
                    if (checkBox != null)
                    {
                        var chkbox = checkBox;
                        chkbox.Checked = false;
                    }
                    else
                    {
                        var button = control as RadioButton;
                        if (button != null)
                        {
                            var rdbtn = button;
                            rdbtn.Checked = false;
                        }
                        else
                        {
                            var picker = control as DateTimePicker;
                            if (picker == null) continue;
                            var dtp = picker;
                            dtp.Value = DateTime.Now;
                        }
                    }
                }
            }
        }
    }
}