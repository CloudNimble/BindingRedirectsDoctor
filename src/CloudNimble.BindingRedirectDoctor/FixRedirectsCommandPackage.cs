//------------------------------------------------------------------------------
// <copyright file="FixRedirectsCommandPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Task = System.Threading.Tasks.Task;

namespace CloudNimble.BindingRedirectDoctor
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuids.guidFixRedirectsCommandPackageString)]
    [ProvideAutoLoad(PackageGuids.guidUIContextString, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideUIContextRule(PackageGuids.guidUIContextString,
        name: "Test auto load",
        expression: "DotConfig",
        termNames: new[] { "DotConfig" },
        termValues: new[] { "HierSingleSelectionName:.config$" })]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class FixRedirectsCommandPackage : AsyncPackage
    {

        public DTE2 _dte;
        public static FixRedirectsCommandPackage Instance;
        private static bool _isProcessing;
        private OleMenuCommandService _commandService;

        /// <summary>
        /// Initializes a new instance of the <see cref="FixRedirectsCommand"/> class.
        /// </summary>
        public FixRedirectsCommandPackage()
        {
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            _dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(_dte);
            Instance = this;

            Logger.Initialize(this, "BindingRedirects Doctor");

            _commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Assumes.Present(_commandService);
            AddCommand(0x0100, (s, e) => { _ = System.Threading.Tasks.Task.Run(() => FixBindingRedirects()); });
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandId"></param>
        /// <param name="invokeHandler"></param>
        /// <param name="beforeQueryStatus"></param>
        private void AddCommand(int commandId, EventHandler invokeHandler)
        {
            var cmdId = new CommandID(PackageGuids.guidFixRedirectsCommandPackageCmdSet, commandId);
            var menuCmd = new OleMenuCommand(invokeHandler, cmdId);
            //menuCmd.BeforeQueryStatus += beforeQueryStatus;
            _commandService.AddCommand(menuCmd);
        }

        /// <summary>
        /// 
        /// </summary>
        private void FixBindingRedirects()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _isProcessing = true;

            var files = ProjectHelpers.GetSelectedItemPaths().Where(c => c.ToLower().EndsWith("web.config") || c.ToLower().EndsWith("app.config"));

            if (!files.Any())
            {
                _dte.StatusBar.Text = "Please select a web.config or app.config file to fix.";
                _isProcessing = false;
                return;
            }

            //var projectFolder = ProjectHelpers.GetRootFolder(ProjectHelpers.GetActiveProject());
            int count = files.Count();

            //RWM: Don't mess with these.
            XNamespace defaultNs = "";
            XNamespace assemblyBindingNs = "urn:schemas-microsoft-com:asm.v1";
            var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

            try
            {
                string text = count == 1 ? " file" : " files";
                _dte.StatusBar.Progress(true, $"Fixing {count} config {text}...", AmountCompleted: 1, Total: count + 1);

                Parallel.For(0, count, options, i =>
                {
                    Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                    var assemblyBindings = new XElement(assemblyBindingNs + "assemblyBinding");
                    var newBindings = new SortedDictionary<string, XElement>();
                    var fullPath = files.ElementAt(i);

                    //RWM: Start by backing up the files.
                    File.Copy(fullPath, fullPath + ".bak", true);
                    Logger.Log($"Backup created for {fullPath}.");

                    //RWM: Load the files.
                    var config = XDocument.Load(fullPath);

                    var oldBindingRoot = config.Root.Descendants().FirstOrDefault(c => c.Name.LocalName == "assemblyBinding");
                    var oldCount = oldBindingRoot.Elements().Count();

                    foreach (var dependentAssembly in oldBindingRoot.Elements().ToList())
                    {
                        var assemblyIdentity = dependentAssembly.Element(assemblyBindingNs + "assemblyIdentity");
                        var bindingRedirect = dependentAssembly.Element(assemblyBindingNs + "bindingRedirect");

                        if (newBindings.ContainsKey(assemblyIdentity.Attribute("name").Value))
                        {
                            Logger.Log($"Reference already exists for {assemblyIdentity.Attribute("name").Value}. Checking version...");
                            //RWM: We've seen this assembly before. Check to see if we can update the version.
                            var newBindingRedirect = newBindings[assemblyIdentity.Attribute("name").Value].Descendants(assemblyBindingNs + "bindingRedirect").First();
                            var oldVersion = Version.Parse(newBindingRedirect.Attribute("newVersion").Value);
                            var newVersion = Version.Parse(bindingRedirect.Attribute("newVersion").Value);

                            if (newVersion > oldVersion)
                            {
                                newBindingRedirect.ReplaceWith(bindingRedirect);
                                Logger.Log($"Version was newer. Binding updated.");
                            }
                            else
                            {
                                Logger.Log($"Version was the same or older. No update needed. Skipping.");
                            }
                        }
                        else
                        {
                            newBindings.Add(assemblyIdentity.Attribute("name").Value, dependentAssembly);
                        }
                    }

                    //RWM: Add the SortedDictionary items to our new assemblyBindingd element.
                    foreach (var binding in newBindings)
                    {
                        assemblyBindings.Add(binding.Value);
                    }

                    //RWM: Fix up the web.config by adding the new assemblyBindings and removing the old one.
                    oldBindingRoot.AddBeforeSelf(assemblyBindings);
                    oldBindingRoot.Remove();

                    //RWM: Save the config file.
                    if (_dte.SourceControl.IsItemUnderSCC(fullPath) && !_dte.SourceControl.IsItemCheckedOut(fullPath))
                    {
                        _dte.SourceControl.CheckOutItem(fullPath);
                    }
                    config.Save(fullPath);

                    Logger.Log($"Update complete. Result: {oldCount} bindings before, {newBindings.Count} after.");
                });
            }
            catch (AggregateException agEx)
            {
                _dte.StatusBar.Progress(false);
                Logger.Log($"Update failed. Exceptions:");
                foreach (var ex in agEx.InnerExceptions)
                {
                    Logger.Log($"Message: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
                }
                _dte.StatusBar.Text = "Operation failed. Please see Output Window for details.";
                _isProcessing = false;

            }
            finally
            {
                _dte.StatusBar.Progress(false);
                _dte.StatusBar.Text = "Operation finished. Please see Output Window for details.";
                _isProcessing = false;
            }

        }

    }

}