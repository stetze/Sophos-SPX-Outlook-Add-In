using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Sophos_SPX_Outlook_Add_In
{
    public partial class MyRibbon
    {
        private const string SpxCategory = "Sophos SPX";
        private const string SpxFlagProperty = "SophosSPXEncrypt";

        // ⭐ Marker: Nur wenn wir selbst Sensitivity gesetzt haben, dürfen wir später zurücksetzen
        private const string SpxSensitivityMarkerProperty = "SophosSPXSensitivitySetByAddin";

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            RefreshUI();
        }

        private Outlook.MailItem GetMail()
        {
            var insp = this.Context as Outlook.Inspector;
            return insp?.CurrentItem as Outlook.MailItem;
        }

        private bool GetState(Outlook.MailItem mail)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty prop = null;

            try
            {
                props = mail.UserProperties;
                prop = props.Find(SpxFlagProperty);

                if (prop != null && prop.Value is bool b)
                    return b;

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (prop != null) Marshal.ReleaseComObject(prop);
                if (props != null) Marshal.ReleaseComObject(props);
            }
        }

        private void RefreshUI()
        {
            Outlook.MailItem mail = null;
            try
            {
                mail = GetMail();
                if (mail == null) return;

                buttonEncrypt.Checked = GetState(mail);
            }
            catch { }
            finally
            {
                if (mail != null) Marshal.ReleaseComObject(mail);
            }
        }

        private void ButtonEncrypt_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MailItem mail = null;

            try
            {
                mail = GetMail();
                if (mail == null) return;

                bool enabled = !GetState(mail);

                UpdateCategory(mail, enabled);
                SetSpxFlag(mail, enabled);
                ApplySophosSensitivityLogic(mail, enabled); // ⭐ NEU: Vertraulich setzen wie Sophos

                // UI state
                buttonEncrypt.Checked = enabled;

                // Persistiere Änderungen (entspricht typischem Add-in Verhalten)
                try { mail.Save(); } catch { }
            }
            catch { }
            finally
            {
                if (mail != null) Marshal.ReleaseComObject(mail);
            }
        }

        private void UpdateCategory(Outlook.MailItem mail, bool enable)
        {
            try
            {
                string categories = mail.Categories ?? string.Empty;
                var list = new System.Collections.Generic.List<string>();

                foreach (var c in categories.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = c.Trim();
                    if (!trimmed.Equals(SpxCategory, StringComparison.OrdinalIgnoreCase))
                        list.Add(trimmed);
                }

                if (enable) list.Add(SpxCategory);

                mail.Categories = string.Join(", ", list);
            }
            catch { }
        }

        private void SetSpxFlag(Outlook.MailItem mail, bool enable)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty prop = null;

            try
            {
                props = mail.UserProperties;
                prop = props.Find(SpxFlagProperty);

                if (enable)
                {
                    if (prop == null)
                        prop = props.Add(SpxFlagProperty, Outlook.OlUserPropertyType.olYesNo, false);

                    prop.Value = true;
                }
                else
                {
                    if (prop != null)
                        prop.Delete();
                }
            }
            catch { }
            finally
            {
                if (prop != null) Marshal.ReleaseComObject(prop);
                if (props != null) Marshal.ReleaseComObject(props);
            }
        }

        /// <summary>
        /// Sophos-like Verhalten:
        /// - Beim Aktivieren: Setzt Sensitivity auf Confidential, aber nur wenn aktuell "Normal"
        /// - Beim Deaktivieren: Setzt nur dann zurück auf Normal, wenn wir es zuvor gesetzt haben
        /// </summary>
        private void ApplySophosSensitivityLogic(Outlook.MailItem mail, bool enable)
        {
            Outlook.UserProperties props = null;
            Outlook.UserProperty marker = null;

            try
            {
                props = mail.UserProperties;
                marker = props.Find(SpxSensitivityMarkerProperty);

                if (enable)
                {
                    // Nur setzen, wenn der Benutzer noch nichts "Strengeres/Anderes" gewählt hat
                    if (mail.Sensitivity == Outlook.OlSensitivity.olNormal)
                    {
                        mail.Sensitivity = Outlook.OlSensitivity.olConfidential;

                        // Marker auf TRUE setzen (oder anlegen)
                        if (marker == null)
                            marker = props.Add(SpxSensitivityMarkerProperty, Outlook.OlUserPropertyType.olYesNo, false);
                        marker.Value = true;
                    }
                    else
                    {
                        // Benutzer hat schon manuell etwas gesetzt → wir respektieren das.
                        // Marker entfernen, damit wir später nichts zurücksetzen.
                        if (marker != null)
                            marker.Delete();
                    }
                }
                else
                {
                    // Nur zurücksetzen, wenn wir es zuvor gesetzt haben
                    bool setByAddin = false;
                    if (marker != null && marker.Value is bool b)
                        setByAddin = b;

                    if (setByAddin)
                    {
                        // Nur zurück auf Normal, wenn es noch Confidential ist (User könnte inzwischen umgestellt haben)
                        if (mail.Sensitivity == Outlook.OlSensitivity.olConfidential)
                            mail.Sensitivity = Outlook.OlSensitivity.olNormal;

                        try { marker.Delete(); } catch { }
                    }
                    else
                    {
                        // Kein Marker => Finger weg von Sensitivity
                        if (marker != null)
                        {
                            try { marker.Delete(); } catch { }
                        }
                    }
                }
            }
            catch { }
            finally
            {
                if (marker != null) Marshal.ReleaseComObject(marker);
                if (props != null) Marshal.ReleaseComObject(props);
            }
        }
    }
}
