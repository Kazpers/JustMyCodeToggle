using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace Kazpers.JustMyCodeToggle
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
    [Guid(PackageGuidString)]
    [ProvideMenuResource(1000, 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class JustMyCodeTogglePackage : AsyncPackage
    {
        /// <summary>
        /// Kazpers.JustMyCodeTogglePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "ad7c95dd-ac83-4811-994b-5f348f2a83dd";
        public static readonly Guid GuidJustMyCodeToggleCommandSet = new Guid("{" + PackageGuidString + "}");
        public static readonly int CmdidJustMyCodeToggle = 0x0100;

        #region Package Members

        private readonly OleMenuCommand _command;

        public JustMyCodeTogglePackage()
        {
            var id = new CommandID(GuidJustMyCodeToggleCommandSet, CmdidJustMyCodeToggle);
            EventHandler invokeHandler = HandleInvokeJustMyCodeToggle;
            EventHandler changeHandler = HandleChangeJustMyCodeToggle;
            EventHandler beforeQueryStatus = HandleBeforeQueryStatusJustMyCodeToggle;
            _command = new OleMenuCommand(invokeHandler, changeHandler, beforeQueryStatus, id);
        }
        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var mcs = (IMenuCommandService)await GetServiceAsync(typeof(IMenuCommandService));
            Assumes.Present(mcs);
            mcs.AddCommand(_command);
        }

        public EnvDTE.DTE ApplicationObject
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return GetService(typeof(EnvDTE._DTE)) as EnvDTE.DTE;
            }
        }

        private void HandleInvokeJustMyCodeToggle(object sender, EventArgs e)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                EnvDTE.Property enableJustMyCode = ApplicationObject.get_Properties("Debugging", "General").Item("EnableJustMyCode");
                if (enableJustMyCode.Value is bool value)
                {
                    enableJustMyCode.Value = !value;
                }
            }
            catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
            {
            }
        }

        private void HandleChangeJustMyCodeToggle(object sender, EventArgs e)
        {
        }

        private void HandleBeforeQueryStatusJustMyCodeToggle(object sender, EventArgs e)
        {
            try
            {
                _command.Supported = true;

                ThreadHelper.ThrowIfNotOnUIThread();
                EnvDTE.Property enableJustMyCode = ApplicationObject.get_Properties("Debugging", "General").Item("EnableJustMyCode");
                if (enableJustMyCode.Value is bool value)
                {
                    _command.Checked = value;
                }

                _command.Enabled = true;
            }
            catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
            {
                _command.Supported = false;
                _command.Enabled = false;
            }
        }
        #endregion
    }
}
