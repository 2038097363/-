using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace MatrixWorkspaceApp
{
    internal static class Program
    {
        private static readonly Dictionary<string, string> EmbeddedFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "matrix_calculator.ps1", "matrix_calculator_embedded.ps1" },
            { "modules.matrix_core.ps1", Path.Combine("modules", "matrix_core.ps1") },
            { "modules.matrix_ui_helpers.ps1", Path.Combine("modules", "matrix_ui_helpers.ps1") },
            { "modules.matrix_dialogs.ps1", Path.Combine("modules", "matrix_dialogs.ps1") },
            { "modules.matrix_form.ps1", Path.Combine("modules", "matrix_form.ps1") },
            { "modules.data_profile.ps1", Path.Combine("modules", "data_profile.ps1") },
            { "modules.data_plotting.ps1", Path.Combine("modules", "data_plotting.ps1") },
            { "modules.data_plotting_ext.ps1", Path.Combine("modules", "data_plotting_ext.ps1") }
        };

        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                string appDirectory = Path.GetDirectoryName(Application.ExecutablePath) ?? AppDomain.CurrentDomain.BaseDirectory;
                string scriptPath = ResolveScriptPath(appDirectory);
                string workingDirectory = Path.GetDirectoryName(scriptPath) ?? appDirectory;

                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + scriptPath + "\"",
                    WorkingDirectory = workingDirectory,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                startInfo.EnvironmentVariables["MATRIX_WORKSPACE_ROOT"] = appDirectory;
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Failed to start Matrix Workspace.\r\n\r\n" + ex.Message,
                    "Matrix Workspace",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string ResolveScriptPath(string appDirectory)
        {
            string looseScriptPath = Path.Combine(appDirectory, "matrix_calculator.ps1");
            string looseModuleDirectory = Path.Combine(appDirectory, "modules");
            if (File.Exists(looseScriptPath) && Directory.Exists(looseModuleDirectory))
            {
                return looseScriptPath;
            }

            return ExtractEmbeddedFiles();
        }
        private static string ExtractEmbeddedFiles()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string targetDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MatrixWorkspaceApp");

            Directory.CreateDirectory(targetDirectory);

            foreach (KeyValuePair<string, string> entry in EmbeddedFiles)
            {
                using (Stream resourceStream = assembly.GetManifestResourceStream(entry.Key))
                {
                    if (resourceStream == null)
                    {
                        throw new InvalidOperationException("Missing embedded resource: " + entry.Key);
                    }

                    string targetPath = Path.Combine(targetDirectory, entry.Value);
                    string targetFolder = Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrEmpty(targetFolder))
                    {
                        Directory.CreateDirectory(targetFolder);
                    }

                    using (var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.Read))
                    {
                        resourceStream.CopyTo(fileStream);
                    }
                }
            }

            return Path.Combine(targetDirectory, "matrix_calculator_embedded.ps1");
        }
    }
}
