using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlookMailItemTaskPane
{
    public class RadioGroup : GroupBox
    {
        public event EventHandler SelectedIndexChanged;
        public int Index
        {
            get { return _mIndex; }
        }
        public RadioButton Button
        {
            get { return _mIndex < 0 ? null : _mButtons[_mIndex]; }
        }

        private readonly List<RadioButton> _mButtons = new List<RadioButton>();
        private int _mIndex = -1;
        protected override void OnVisibleChanged(EventArgs e)
        {
            // Hijack this event to initialize the control
            if (_mButtons.Count == 0)
            {
                // Build list of radio buttons
                foreach (var btn in Controls.OfType<RadioButton>())
                {
                    _mButtons.Add(btn);
                    btn.CheckedChanged += ButtonCheckChanged;
                }
                // Sort list by tab index so Index property is meaningful
                _mButtons.Sort(SortByTabIndex);
                // Generate initial Index and SelectedIndexChanged event
                for (_mIndex = _mButtons.Count - 1; _mIndex > 0; _mIndex--)
                    if (_mButtons[_mIndex].Checked)
                    {
                        if (SelectedIndexChanged != null) SelectedIndexChanged.Invoke(this, EventArgs.Empty);
                        break;
                    }
            }
            base.OnVisibleChanged(e);
        }
        private static int SortByTabIndex(RadioButton btn1, RadioButton btn2)
        {
            // Sort helper
            return btn1.TabIndex < btn2.TabIndex ? -1 : 1;
        }
        private void ButtonCheckChanged(object sender, EventArgs e)
        {
            // Generate SelectedIndexChanged event
            var btn = (RadioButton)sender;
            if (!btn.Checked) return;
            for (_mIndex = 0; _mIndex < _mButtons.Count; ++_mIndex)
                if (ReferenceEquals(btn, _mButtons[_mIndex]))
                {
                    if (SelectedIndexChanged != null) SelectedIndexChanged.Invoke(this, EventArgs.Empty);
                    return;
                }
            _mIndex = -1;
        }
        protected override void Dispose(bool disposing)
        {
            // Release references to radio buttons
            if (disposing) _mButtons.Clear();
            base.Dispose(disposing);
        }
    }
}
