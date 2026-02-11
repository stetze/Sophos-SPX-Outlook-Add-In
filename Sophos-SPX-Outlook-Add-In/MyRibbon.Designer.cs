using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace Sophos_SPX_Outlook_Add_In
{
    partial class MyRibbon : RibbonBase
    {
        private RibbonTab tab;
        internal RibbonToggleButton buttonEncrypt;

        public MyRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.tab = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.buttonEncrypt = this.Factory.CreateRibbonToggleButton();
            this.tab.SuspendLayout();
            this.group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab
            // 
            this.tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab.ControlId.OfficeId = "TabNewMailMessage";
            this.tab.Groups.Add(this.group);
            this.tab.Label = "TabNewMailMessage";
            this.tab.Name = "tab";
            // 
            // group
            // 
            this.group.Items.Add(this.buttonEncrypt);
            this.group.Label = "Sophos";
            this.group.Name = "group";
            // 
            // buttonEncrypt
            // 
            this.buttonEncrypt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonEncrypt.Label = "Verschlüsseln";
            this.buttonEncrypt.Name = "buttonEncrypt";
            this.buttonEncrypt.OfficeImageId = "EncryptMessage";
            this.buttonEncrypt.ShowImage = true;
            this.buttonEncrypt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonEncrypt_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.ResumeLayout(false);

        }

        public RibbonGroup group;
    }
}
