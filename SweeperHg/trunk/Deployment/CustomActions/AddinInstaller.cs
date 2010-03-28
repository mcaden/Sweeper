namespace CustomActions
{
    using System;
    using System.Collections;
    using System.ComponentModel;
    using System.Configuration.Install;
    using System.Diagnostics;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Installer Custom Actions for the Addin
    /// </summary>
    [RunInstaller(true)]
    public partial class AddinInstaller : Installer
    {
        /// <summary>
        /// Namespace used in the .addin configuration file.
        /// </summary>         
        private const string ExtNameSpace =
          "http://schemas.microsoft.com/AutomationExtensibility";

        /// <summary>
        /// Initializes a new instance of the AddinInstaller class.
        /// </summary>
        public AddinInstaller()
            : base()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Overrides Installer.Install,
        /// which will be executed during install process.
        /// </summary>
        /// <param name="savedState">The saved state.</param>
        public override void Install(IDictionary savedState)
        {
            // Uncomment the following line, recompile the setup
            // project and run the setup executable if you want
            // to debug into this custom action.

            ////Debugger.Break();

            base.Install(savedState);

            // Parameters required to pass in from installer
            string productName = this.Context.Parameters["ProductName"];
            string assemblyName = this.Context.Parameters["AssemblyName"];

            // Setup .addin path and assembly path
            string addinTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2008\Addins");
            string assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string addinControlFileName = assemblyName + ".Addin";
            string addinAssemblyFileName = assemblyName + ".dll";

            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(addinTargetPath);
                if (!dirInfo.Exists)
                {
                    dirInfo.Create();
                }

                string sourceFile = Path.Combine(assemblyPath, addinControlFileName);
                XmlDocument doc = new XmlDocument();
                doc.Load(sourceFile);
                XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
                xnm.AddNamespace("def", ExtNameSpace);

                // Update Addin/Assembly node
                XmlNode node = doc.SelectSingleNode("/def:Extensibility/def:Addin/def:Assembly", xnm);
                if (node != null)
                {
                    node.InnerText = Path.Combine(assemblyPath, addinAssemblyFileName);
                }

                // Update ToolsOptionsPage/Assembly node
                node = doc.SelectSingleNode("/def:Extensibility/def:ToolsOptionsPage/def:Category/def:SubCategory/def:Assembly", xnm);
                if (node != null)
                {
                    node.InnerText = Path.Combine(assemblyPath, addinAssemblyFileName);
                }

                doc.Save(sourceFile);

                string targetFile = Path.Combine(addinTargetPath, addinControlFileName);
                File.Copy(sourceFile, targetFile, true);

                // Save AddinPath to be used in Uninstall or Rollback
                savedState.Add("AddinPath", targetFile);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Overrides Installer.Rollback, which will be executed during rollback process.
        /// </summary>
        /// <param name="savedState">The saved state.</param>
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);

            try
            {
                string fileName = (string)savedState["AddinPath"];
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Overrides Installer.Uninstall, which will be executed during uninstall process.
        /// </summary>
        /// <param name="savedState">The saved state.</param>
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            try
            {
                string fileName = (string)savedState["AddinPath"];
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}