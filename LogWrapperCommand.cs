using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;

using EnvDTE;

using Microsoft.VisualStudio.Shell;

using Task = System.Threading.Tasks.Task;

namespace LogWrapperCommand
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class LogWrapperCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("459d275f-56f2-4e85-84a5-077de78a8394");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private DTE dteInstance;

        /// <summary>
        /// Initializes a new instance of the <see cref="LogWrapperCommand" /> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// <param name="dteClass">The DTE class.</param>
        /// <exception cref="ArgumentNullException">
        /// package
        /// or
        /// commandService
        /// </exception>
        private LogWrapperCommand(AsyncPackage package, OleMenuCommandService commandService, DTE dteClass)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));

            this.dteInstance = dteClass;

            // Register command
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static LogWrapperCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE;

            Instance = new LogWrapperCommand(package, commandService, dte);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var activeDocument = this.dteInstance.ActiveDocument as EnvDTE.Document;

            //var methods = SearchCollection<CodeFunction>(activeDocument.ProjectItem.FileCodeModel.CodeElements);
            //foreach(var singleMethod in methods)
            //{
            //    var code = InnerText((CodeElement)singleMethod);

            //    System.Diagnostics.Trace.WriteLine(code);
            //}



            var selection = activeDocument.Selection as TextSelection;
            var funcAtCursor = selection.ActivePoint.CodeElement[vsCMElement.vsCMElementFunction] as CodeElement;
            if (funcAtCursor == null)
            {
                // Not inside function
                return;
            }

            var funcCode = InnerText(funcAtCursor);
            //if (funcCode.Contains()) ;

            AddTraceLogs(funcAtCursor);
        }

        /// <summary>
        /// Adds the trace logs to the first and last lines of function
        /// </summary>
        /// <param name="function">The function.</param>
        private void AddTraceLogs(CodeElement function)
        {
            TextSelection txtSel = (TextSelection)function.DTE.ActiveDocument.Selection;

            var editPoint = function.GetStartPoint().CreateEditPoint();

            //
            // Prolog
            //
            var lineOfCode = GetLineText(editPoint);

            while (lineOfCode.Trim() != "{")
            {
                editPoint.LineDown();

                lineOfCode = GetLineText(editPoint);
            }

            // Scroll to the beginning of the line
            //editPoint.s StartOfLine();

            // Just a after '{' symbol - one line down
            editPoint.LineDown();

            // Move 'selection' to point above
            txtSel.MoveToPoint(editPoint);

            // Add blank line
            txtSel.NewLine(1);

            // Return selection to line above
            txtSel.MoveToPoint(editPoint);

            var prologText = ((LogWrapperCommandPackage)this.package).PrologText;
            prologText = Substitute(prologText, function);
            txtSel.Insert(prologText);

            // Epilog
            editPoint = function.GetEndPoint().CreateEditPoint();
            
            lineOfCode = GetLineText(editPoint);

            while (lineOfCode.Trim() != "}")
            {
                editPoint.LineDown();

                lineOfCode = GetLineText(editPoint);
            }

            // Move the point just before the '}' symbol
            editPoint.CharLeft(1);

            // Move selection there
            txtSel.MoveToPoint(editPoint);

            // Add new line
            txtSel.NewLine(1);

            // Move selection there
            txtSel.MoveToPoint(editPoint);

            var epilogText = ((LogWrapperCommandPackage)this.package).EpilogText;
            
            epilogText = Substitute(epilogText, function);

            // Insert text
            txtSel.Insert(epilogText);
        }

        /// <summary>
        /// Substitutes the specified text
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="function">The function.</param>
        /// <returns></returns>
        private string Substitute(string text, CodeElement function)
        {
            if (text.Contains("{functionName}"))
            {
                text = text.Replace("{functionName}", function.Name);
            }

            return text;
        }

        /// <summary>
        /// Gets the text in line where provided <paramref name="point"/> stand
        /// </summary>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        public static string GetLineText(EditPoint point)
        {
            point.StartOfLine();
            var start = point.CreateEditPoint();

            point.EndOfLine();
            var end = point.CreateEditPoint();

            return start.GetText(end);
        }


        /// <summary>
        /// Searches the provided <paramref name="codeElement"/> for elements of type <see cref="T"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="codeElement">The code element.</param>
        /// <returns>Collection of <see cref="T"/> elements if any</returns>
        private IEnumerable<T> SearchSingle<T>(CodeElement codeElement)
        {
            var coll = new List<T>();

            if (codeElement is T)
            {
                coll.Add((T)codeElement);
            }
            
            var objCodeNamespace = codeElement as CodeNamespace;
            if (objCodeNamespace != null)
            {
                var result = SearchCollection<T>(objCodeNamespace.Members);
                coll.AddRange(result);
            }

            var objCodeType = codeElement as CodeType;
            if (objCodeType != null)
            {
                var result = SearchCollection<T>(objCodeType.Members);
                coll.AddRange(result);
            }

            return coll;
        }

        /// <summary>
        /// Searches the for elements of type <see cref="T"/> in provided <paramref name="codeElementsCollection"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="codeElementsCollection">The code elements collection.</param>
        /// <returns></returns>
        private IEnumerable<T> SearchCollection<T>(CodeElements codeElementsCollection)
        {
            var coll = new List<T>();

            foreach(CodeElement singleElement in codeElementsCollection)
            {
                var result = SearchSingle<T>(singleElement);
                if ((result != null) && (result.Any() == true))
                {
                    coll.AddRange(result);
                }
            }

            return coll;
        }

        /// <summary>
        /// Function will return inner text of provided <paramref name="codeElement"/>
        /// </summary>
        /// <param name="codeElement">The code element.</param>
        /// <returns></returns>
        private string InnerText(CodeElement codeElement)
        {
            return codeElement.GetStartPoint().CreateEditPoint().GetText(codeElement.EndPoint);
        }
    }
}
