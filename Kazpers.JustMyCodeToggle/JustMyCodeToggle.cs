using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace Kazpers.JustMyCodeToggle
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class JustMyCodeToggle
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9a0e9fc4-d857-4e76-ad95-95a78432a1d0");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="JustMyCodeToggle"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private JustMyCodeToggle(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(OnInvoke, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static JustMyCodeToggle Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider => _package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in JustMyCodeToggle's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new JustMyCodeToggle(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void OnInvoke(object sender, EventArgs e)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var package = _package as JustMyCodeTogglePackage;
                var applicationObject = package?.ApplicationObject;
                Property enableJustMyCode = applicationObject?.get_Properties("Debugging", "General").Item("EnableJustMyCode");
                if (enableJustMyCode?.Value is bool value)
                {
                    enableJustMyCode.Value = !value;
                }
            }
            catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
            {
            }
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            if (!(sender is OleMenuCommand command)) return;

            try
            {
                command.Supported = true;

                ThreadHelper.ThrowIfNotOnUIThread();
                var package = _package as JustMyCodeTogglePackage;
                var applicationObject = package?.ApplicationObject;
                Property enableJustMyCode = applicationObject?.get_Properties("Debugging", "General").Item("EnableJustMyCode");
                if (enableJustMyCode?.Value is bool value)
                {
                    command.Checked = value;
                }

                command.Enabled = true;
            }
            catch (Exception ex) when (!ErrorHandler.IsCriticalException(ex))
            {
                command.Supported = false;
                command.Enabled = false;
            }
        }
    }
}
