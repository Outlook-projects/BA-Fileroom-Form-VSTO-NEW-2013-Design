using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace OutlookMailItemTaskPane
{
    public class InspectorWrapper
    {
        private Outlook.Inspector _inspector;

        public InspectorWrapper(Outlook.Inspector inspector)
        {
            _inspector = inspector;
            ((Outlook.InspectorEvents_Event)_inspector).Close +=
                InspectorWrapper_Close;

            CustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new OutlookMailItemTaskPane(), "BA Fileroom Form", _inspector);
            CustomTaskPane.Width = 440;
            CustomTaskPane.VisibleChanged += TaskPane_VisibleChanged;
            CustomTaskPane.Control.Focus();
        }

        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons[_inspector].ManageTaskPaneRibbon.toggleButton1.Checked =
                CustomTaskPane.Visible;
        }

        private void InspectorWrapper_Close()
        {
            if (CustomTaskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(CustomTaskPane);
            }

            CustomTaskPane = null;
            Globals.ThisAddIn.InspectorWrappers.Remove(_inspector);
            ((Outlook.InspectorEvents_Event)_inspector).Close -=
                InspectorWrapper_Close;
            _inspector = null;

            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        }

        public CustomTaskPane CustomTaskPane { get; private set; }
    }

    public partial class ThisAddIn
    {
        private Dictionary<Outlook.Inspector, InspectorWrapper> _inspectorWrappersValue =
            new Dictionary<Outlook.Inspector, InspectorWrapper>();
        private Outlook.Inspectors _inspectors;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector +=
                Inspectors_NewInspector;

            foreach (Outlook.Inspector inspector in _inspectors)
            {
                Inspectors_NewInspector(inspector);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _inspectors.NewInspector -=
                Inspectors_NewInspector;
            _inspectors = null;
            _inspectorWrappersValue = null;
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem)
            {
                _inspectorWrappersValue.Add(inspector, new InspectorWrapper(inspector));
            }
        }

        public Dictionary<Outlook.Inspector, InspectorWrapper> InspectorWrappers
        {
            get
            {
                return _inspectorWrappersValue;
            }
        }

        #region VSTO generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor. 
        /// </summary> 
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}