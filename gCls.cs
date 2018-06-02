using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace ePubFix
{
    public static class gCls
    {
        public static int lic_totalpage = 0;
        public static int lic_runpage = 0;

        public static Aspose.Pdf.License aspose_license = new Aspose.Pdf.License();
        public static string aspose_key = "";
        public static string pdfex_key = "";

        public static void update_path_var()
        {
            try
            {
                string gv = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine);
                string xPath = Application.StartupPath + "\\pdfex";
                if (gv != "")
                {
                    string[] allv = gv.Split(';');
                    bool pathFound = false;
                    foreach (string s in allv)
                    {
                        if (s == xPath)
                        {
                            pathFound = true;
                        }
                    }
                    if (pathFound == false)
                    {
                        Environment.SetEnvironmentVariable("PATH", gv + ";" + xPath, EnvironmentVariableTarget.Machine);
                    }

                }


            }
            catch (Exception erd)
            {
                show_error(erd.Message.ToString() + "\n\nRun as administrator to registry access");
                return;
            }

        }

        public static string HexConverter(System.Drawing.Color c)
        {
            String rtn = String.Empty;
            try
            {
                rtn = "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
            }
            catch (Exception ex)
            {
                //doing nothing
            }

            return rtn;
        }

        public static string Text2HexaConversion(string text)
        {
            char[] chars = text.ToCharArray();
            StringBuilder result = new StringBuilder(text.Length + (int)(text.Length * 0.1));

            foreach (char c in chars)
            {
                int value = (int)(c);
                string fhvalue = value.ToString("X");
                int dvalue = int.Parse(fhvalue, System.Globalization.NumberStyles.HexNumber);
                string hvalue = dvalue.ToString();

                //if (hvalue.Length == 2) {
                //    hvalue = "00" + hvalue;
                //}
                //else if (hvalue.Length == 3) {
                //    hvalue = "0" + hvalue;
                //}

                if (value > 127)
                    result.AppendFormat("&#{0};", hvalue);
                else
                    result.Append(c);
            }

            return result.ToString();
        }

        public static void setdate_update(string dateStr)
        {
            try
            {

                System.Diagnostics.ProcessStartInfo psInfo = new System.Diagnostics.ProcessStartInfo();
                psInfo.CreateNoWindow = true;
                psInfo.UseShellExecute = false;
                psInfo.RedirectStandardOutput = true;
                psInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                psInfo.FileName = "cmd";
                psInfo.Arguments = "CMD /c DATE " + dateStr;
                using (System.Diagnostics.Process exeProcess = System.Diagnostics.Process.Start(psInfo))
                {
                    exeProcess.WaitForExit();
                }

            }
            catch { }

        }

        public static void show_error(string msg)
        {
            try
            {

                ComponentFactory.Krypton.Toolkit.KryptonMessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch { }

        }
        public static void show_message(string msg)
        {
            try
            {
                ComponentFactory.Krypton.Toolkit.KryptonMessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch { }

        }

       


       

        public static void getlicstring()
        {
            try
            {
                //enter aspose license key here
                aspose_key = "";

                //enter pdfex license key here
                pdfex_key = "";


            }
            catch (Exception erd)
            {
                show_error("Unable to get license information");
                Application.Exit();
            }

        }

        public static void get_image_size(string imgPath, ref int iWidth, ref int iHeight)
        {

            try
            {
                System.Drawing.Image imgBox = System.Drawing.Image.FromFile(imgPath);
                iWidth = imgBox.Size.Width;
                iHeight = imgBox.Size.Height;
                imgBox.Dispose();
            }
            catch { }

        }

    }
}

