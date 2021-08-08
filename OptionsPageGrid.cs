using System.ComponentModel;

using Microsoft.VisualStudio.Shell;

namespace LogWrapperCommand
{
    public class OptionsPageGrid : DialogPage
    {
        public OptionsPageGrid()
        {
            this.PrologText = "prolog";
            this.EpilogText = "epilog";
        }

        [Category("LogWrapper")]
        [DisplayName("Prolog text")]
        [Description("Text that should be inserted as 'prolog'")]    
        public string PrologText
        {
            get;
            set;
        }

        [Category("LogWrapper")]
        [DisplayName("Epilog text")]
        [Description("Text that should be inserted as 'epilog'")]

        public string EpilogText
        {
            get;
            set;
        }
    }
}
