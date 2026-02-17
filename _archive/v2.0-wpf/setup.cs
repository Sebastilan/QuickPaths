using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows.Forms;

class QuickPathsSetup
{
    [STAThread]
    static void Main()
    {
        string installDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "QuickPaths");
        string exeUrl = "https://github.com/Sebastilan/QuickPaths/releases/latest/download/QuickPaths.exe";
        string exePath = Path.Combine(installDir, "QuickPaths.exe");

        try
        {
            // Create install dir
            Directory.CreateDirectory(installDir);

            // Download exe
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (var wc = new WebClient())
            {
                wc.DownloadFile(exeUrl, exePath);
            }

            // Run install
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "--install",
                UseShellExecute = true
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Installation failed:\n" + ex.Message,
                "QuickPaths Setup",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
