using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Sophos_SPX_Outlook_Add_In
{
    public partial class ThisAddIn
    {
        private const string SpxFlagProperty = "SophosSPXEncrypt";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Application.ItemSend -= Application_ItemSend;
            }
            catch { }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            Outlook.MailItem mail = null;
            Outlook.UserProperties props = null;
            Outlook.UserProperty prop = null;
            Outlook.PropertyAccessor pa = null;

            try
            {
                mail = item as Outlook.MailItem;
                if (mail == null)
                    return;

                bool encrypt = false;

                props = mail.UserProperties;
                prop = props.Find(SpxFlagProperty);
                if (prop != null && prop.Value is bool b)
                    encrypt = b;

                const string SPX_HEADER =
                    "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/x-sophos-spx-encrypt";

                pa = mail.PropertyAccessor;

                if (encrypt)
                {
                    pa.SetProperty(SPX_HEADER, "yes");
                }
                else
                {
                    try { pa.DeleteProperty(SPX_HEADER); } catch { }
                }

                mail.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SPX error: " + ex.Message);
            }
            finally
            {
                if (pa != null) Marshal.ReleaseComObject(pa);
                if (prop != null) Marshal.ReleaseComObject(prop);
                if (props != null) Marshal.ReleaseComObject(props);
                if (mail != null) Marshal.ReleaseComObject(mail);
            }
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}
