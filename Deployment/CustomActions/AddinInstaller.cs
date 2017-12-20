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
        private const string ExtNameSpace = "http://schemas.microsoft.com/AutomationExtensibility";

        /// <summary>
        /// Install state for visual studio 2008.
        /// </summary>
        private const string VisualStudio2008State = "AddinPath2008";

        /// <summary>
        /// Install state for visual studio 2010.
        /// </summary>
        private const string VisualStudio2010State = "AddinPath2010";

        /// <summary>
        /// Install state for visual studio 2010.
        /// </summary>
        private const string VisualStudio2012State = "AddinPath2012";

        /// <summary>
        /// Install state for visual studio 2010.
        /// </summary>
        private const string VisualStudio2013State = "AddinPath2013";

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
            //// Debugger.Break();
            base.Install(savedState);

            // Parameters required to pass in from installer
            string productName = this.Context.Parameters["ProductName"];
            string assemblyName = this.Context.Parameters["AssemblyName"];

            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2008")))
            {
                // Setup .addin path and assembly path
                string addinTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2008\Addins");

                SetupAddin(savedState, assemblyName, addinTargetPath, "9.0", VisualStudio2008State);
            }

            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2010")))
            {
                // Setup .addin path and assembly path
                string addinTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2010\Addins");

                SetupAddin(savedState, assemblyName, addinTargetPath, "10.0", VisualStudio2010State);
            }

            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2012")))
            {
                // Setup .addin path and assembly path
                string addinTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2012\Addins");

                SetupAddin(savedState, assemblyName, addinTargetPath, "11.0", VisualStudio2012State);
            }

            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2013")))
            {
                // Setup .addin path and assembly path
                string addinTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), @"Visual Studio 2013\Addins");

                SetupAddin(savedState, assemblyName, addinTargetPath, "12.0", VisualStudio2013State);
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
                string fileNameVS2008 = (string)savedState[VisualStudio2008State];
                if (File.Exists(fileNameVS2008))
                {
                    File.Delete(fileNameVS2008);
                }

                string fileNameVS2010 = (string)savedState[VisualStudio2010State];
                if (File.Exists(fileNameVS2010))
                {
                    File.Delete(fileNameVS2010);
                }

                string fileNameVS2012 = (string)savedState[VisualStudio2012State];
                if (File.Exists(fileNameVS2012))
                {
                    File.Delete(fileNameVS2012);
                }

                string fileNameVS2013 = (string)savedState[VisualStudio2013State];
                if (File.Exists(fileNameVS2013))
                {
                    File.Delete(fileNameVS2013);
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
                string fileNameVS2008 = (string)savedState[VisualStudio2008State];
                if (File.Exists(fileNameVS2008))
                {
                    File.Delete(fileNameVS2008);
                }

                string fileNameVS2010 = (string)savedState[VisualStudio2010State];
                if (File.Exists(fileNameVS2010))
                {
                    File.Delete(fileNameVS2010);
                }

                string fileNameVS2012 = (string)savedState[VisualStudio2012State];
                if (File.Exists(fileNameVS2012))
                {
                    File.Delete(fileNameVS2012);
                }

                string fileNameVS2013 = (string)savedState[VisualStudio2013State];
                if (File.Exists(fileNameVS2013))
                {
                    File.Delete(fileNameVS2013);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Sets up the addin for a version of Visual Studio
        /// </summary>
        /// <param name="savedState">The installer state</param>
        /// <param name="assemblyName">The name of the assembly</param>
        /// <param name="addinTargetPath">The path to the assembly</param>
        /// <param name="version">The version number of visual studio.  9.0 is VS 2008, 10.0 is VS 2010</param>
        /// <param name="state">The name of the saved state.</param>
        private static void SetupAddin(IDictionary savedState, string assemblyName, string addinTargetPath, string version, string state)
        {
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

                // Update Addin/Assembly node
                XmlNodeList versionNodes = doc.SelectNodes("/def:Extensibility/def:HostApplication/def:Version", xnm);
                if (versionNodes != null)
                {
                    for (int i = 0; i < versionNodes.Count; i++)
                    {
                        if (versionNodes[i] != null)
                        {
                            versionNodes[i].InnerText = version;
                        }
                    }
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
                savedState.Add(state, targetFile);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}