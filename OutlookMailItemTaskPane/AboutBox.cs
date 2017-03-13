using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace OutlookMailItemTaskPane
{
    public partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();

            //  Initialize the AboutBox to display the product information from the assembly information.
            //  Change assembly information settings for your application through either:
            //  - Project->Properties->Application->Assembly Information
            //  - AssemblyInfo.cs
            Text = $@"About {AssemblyTitle}";
            label1.Text = AssemblyProduct;
            label2.Text = $@"Version {AssemblyVersion}";
            label3.Text = AssemblyCopyright;
            label4.Text = AssemblyCompany;
            textBox1.Text = AssemblyDescription;
        }

        public sealed override string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        #region Assembly Attribute Accessors

        private static string AssemblyTitle
        {
            get
            {
                // Get all Title attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                // If there is at least one Title attribute
                if (attributes.Length <= 0)
                    return Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
                // Select the first one
                var titleAttribute = (AssemblyTitleAttribute)attributes[0];
                // If it is not an empty string, return it
                return titleAttribute.Title != "" ? titleAttribute.Title : Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
                // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
            }
        }

        private static string AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        private static string AssemblyDescription
        {
            get
            {
                // Get all Description attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                // If there aren't any Description attributes, return an empty string
                return attributes.Length == 0 ? "" : ((AssemblyDescriptionAttribute)attributes[0]).Description;
                // If there is a Description attribute, return its value
            }
        }

        private static string AssemblyProduct
        {
            get
            {
                // Get all Product attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                // If there aren't any Product attributes, return an empty string
                return attributes.Length == 0 ? "" : ((AssemblyProductAttribute)attributes[0]).Product;
                // If there is a Product attribute, return its value
            }
        }

        private static string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                // If there aren't any Copyright attributes, return an empty string
                return attributes.Length == 0 ? "" : ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
                // If there is a Copyright attribute, return its value
            }
        }

        private static string AssemblyCompany
        {
            get
            {
                // Get all Company attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                // If there aren't any Company attributes, return an empty string
                return attributes.Length == 0 ? "" : ((AssemblyCompanyAttribute)attributes[0]).Company;
                // If there is a Company attribute, return its value
            }
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
