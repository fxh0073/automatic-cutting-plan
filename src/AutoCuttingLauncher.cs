using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class AutoCuttingLauncher
{
    private static readonly string AppTitle = "\u81ea\u52a8\u4e0b\u6599\u7a0b\u5e8f";

    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length >= 3 && string.Equals(args[0], "--headless", StringComparison.OrdinalIgnoreCase))
        {
            RunHeadless(args);
            return;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainWindow());
    }

    private static void RunHeadless(string[] args)
    {
        string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last_run.log");
        try
        {
            string result = RunEngine(args[1], args[2], 6000, 3, 5, 2000);
            File.WriteAllText(logPath, "OK\r\n" + result);
            Environment.ExitCode = 0;
        }
        catch (Exception error)
        {
            File.WriteAllText(logPath, "ERROR\r\n" + error);
            Environment.ExitCode = 1;
        }
    }

    private sealed class MainWindow : Form
    {
        private readonly TextBox inputBox;
        private readonly NumericUpDown stockBox;
        private readonly NumericUpDown kerfBox;
        private readonly NumericUpDown stepBox;
        private readonly Button browseButton;
        private readonly Button generateButton;
        private readonly Label statusLabel;

        internal MainWindow()
        {
            Text = AppTitle;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(680, 315);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;
            ShowInTaskbar = true;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label title = new Label();
            title.Text = "\u81ea\u52a8\u4e0b\u6599\u7a0b\u5e8f";
            title.Font = new Font("Microsoft YaHei UI", 17F, FontStyle.Bold, GraphicsUnit.Point);
            title.AutoSize = true;
            title.Location = new Point(24, 18);
            Controls.Add(title);

            AddLabel("\u8f93\u5165\u5de5\u827a\u5361\uff1a", 24, 76);
            inputBox = new TextBox();
            inputBox.Location = new Point(142, 72);
            inputBox.Size = new Size(418, 27);
            Controls.Add(inputBox);

            browseButton = new Button();
            browseButton.Text = "\u9009\u62e9\u6587\u4ef6";
            browseButton.Location = new Point(572, 70);
            browseButton.Size = new Size(84, 30);
            browseButton.Click += BrowseButtonClick;
            Controls.Add(browseButton);

            AddLabel("\u539f\u6599\u5b9a\u5c3a(mm)\uff1a", 24, 124);
            stockBox = AddNumberBox(142, 120, 6000, 1, 100000, 0);

            AddLabel("\u952f\u7f1d(mm)\uff1a", 285, 124);
            kerfBox = AddNumberBox(382, 120, 3, 0, 100, 2);

            AddLabel("\u53d6\u6574\u6b65\u957f(mm)\uff1a", 24, 171);
            stepBox = AddNumberBox(142, 167, 5, 1, 1000, 0);

            Label hint = new Label();
            hint.Text = "\u7ed3\u679c\u5c06\u4fdd\u5b58\u5728\u8f93\u5165\u5de5\u827a\u5361\u6240\u5728\u6587\u4ef6\u5939\uff0c\u4e0d\u8986\u76d6\u539f\u6587\u4ef6\u3002";
            hint.AutoSize = true;
            hint.ForeColor = Color.DimGray;
            hint.Location = new Point(24, 214);
            Controls.Add(hint);

            statusLabel = new Label();
            statusLabel.Text = "\u8bf7\u9009\u62e9\u4e0b\u6599\u5de5\u827a\u5361\u3002";
            statusLabel.AutoSize = true;
            statusLabel.ForeColor = Color.FromArgb(31, 78, 121);
            statusLabel.Location = new Point(24, 267);
            Controls.Add(statusLabel);

            generateButton = new Button();
            generateButton.Text = "\u751f\u6210\u4e0b\u6599\u65b9\u6848";
            generateButton.Location = new Point(422, 252);
            generateButton.Size = new Size(126, 38);
            generateButton.Click += GenerateButtonClick;
            Controls.Add(generateButton);

            Button closeButton = new Button();
            closeButton.Text = "\u5173\u95ed";
            closeButton.Location = new Point(562, 252);
            closeButton.Size = new Size(94, 38);
            closeButton.Click += delegate { Close(); };
            Controls.Add(closeButton);

            AcceptButton = generateButton;
            CancelButton = closeButton;
        }

        private void AddLabel(string text, int x, int y)
        {
            Label label = new Label();
            label.Text = text;
            label.AutoSize = true;
            label.Location = new Point(x, y);
            Controls.Add(label);
        }

        private NumericUpDown AddNumberBox(int x, int y, decimal value, decimal minimum, decimal maximum, int decimals)
        {
            NumericUpDown box = new NumericUpDown();
            box.Location = new Point(x, y);
            box.Size = new Size(112, 27);
            box.Minimum = minimum;
            box.Maximum = maximum;
            box.Value = value;
            box.DecimalPlaces = decimals;
            box.ThousandsSeparator = true;
            Controls.Add(box);
            return box;
        }

        private void BrowseButtonClick(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel \u5de5\u4f5c\u7c3f (*.xlsx)|*.xlsx|\u6240\u6709\u6587\u4ef6 (*.*)|*.*";
                dialog.Title = "\u9009\u62e9\u4e0b\u6599\u5de5\u827a\u5361";
                dialog.Multiselect = false;
                dialog.CheckFileExists = true;
                if (!string.IsNullOrWhiteSpace(inputBox.Text) && File.Exists(inputBox.Text))
                    dialog.InitialDirectory = Path.GetDirectoryName(inputBox.Text);
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    inputBox.Text = dialog.FileName;
                    statusLabel.Text = "\u5df2\u9009\u62e9\uff1a" + Path.GetFileName(dialog.FileName);
                }
            }
        }

        private async void GenerateButtonClick(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(inputBox.Text) || !File.Exists(inputBox.Text))
            {
                MessageBox.Show(this, "\u8bf7\u5148\u9009\u62e9\u6709\u6548\u7684 Excel \u4e0b\u6599\u5de5\u827a\u5361\u3002", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sourcePath = Path.GetFullPath(inputBox.Text);
            string outputPath = Path.Combine(Path.GetDirectoryName(sourcePath), "auto_cutting_plan_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            decimal stock = stockBox.Value;
            decimal kerf = kerfBox.Value;
            decimal step = stepBox.Value;
            SetBusy(true, "\u6b63\u5728\u8ba1\u7b97\u5e76\u751f\u6210 Excel\uff0c\u8bf7\u7a0d\u5019\u2026");

            try
            {
                string result = await Task.Run(delegate
                {
                    return RunEngine(sourcePath, outputPath, stock, kerf, step, 2000);
                });
                statusLabel.Text = "\u751f\u6210\u5b8c\u6210\uff1a" + Path.GetFileName(outputPath);
                MessageBox.Show(this, "\u4e0b\u6599\u65b9\u6848\u5df2\u751f\u6210\uff1a\r\n" + outputPath, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception error)
            {
                string errorLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "auto_cutting_error.log");
                try { File.WriteAllText(errorLog, error.ToString()); } catch { }
                statusLabel.Text = "\u751f\u6210\u5931\u8d25\uff0c\u8bf7\u67e5\u770b auto_cutting_error.log\u3002";
                MessageBox.Show(this, error.Message + "\r\n\r\n\u9519\u8bef\u65e5\u5fd7\uff1a" + errorLog, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false, statusLabel.Text);
            }
        }

        private void SetBusy(bool busy, string status)
        {
            generateButton.Enabled = !busy;
            browseButton.Enabled = !busy;
            inputBox.Enabled = !busy;
            stockBox.Enabled = !busy;
            kerfBox.Enabled = !busy;
            stepBox.Enabled = !busy;
            UseWaitCursor = busy;
            statusLabel.Text = status;
        }
    }

    private static string RunEngine(string sourcePath, string outputPath, decimal stockLength, decimal kerf, decimal lengthStep, int candidates)
    {
        string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string portableEngine = Path.Combine(appDirectory, "engine", "build_cutting_plan.mjs");
        string sourceEngine = Path.GetFullPath(Path.Combine(appDirectory, "..", "src", "build_cutting_plan.mjs"));
        string enginePath = File.Exists(portableEngine) ? portableEngine : sourceEngine;
        if (!File.Exists(enginePath)) throw new FileNotFoundException("Cutting engine not found.", enginePath);

        string nodePath = FindNodeRuntime(appDirectory);
        string arguments = Quote(enginePath)
            + " --input " + Quote(sourcePath)
            + " --output " + Quote(outputPath)
            + " --stock-length " + stockLength.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " --kerf " + kerf.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " --length-step " + lengthStep.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " --candidates " + candidates.ToString(System.Globalization.CultureInfo.InvariantCulture);

        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = nodePath;
        startInfo.Arguments = arguments;
        startInfo.WorkingDirectory = Path.GetDirectoryName(enginePath);
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        startInfo.StandardOutputEncoding = Encoding.UTF8;
        startInfo.StandardErrorEncoding = Encoding.UTF8;

        using (Process process = Process.Start(startInfo))
        {
            if (process == null) throw new InvalidOperationException("Failed to start the cutting engine.");
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                string message = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
                throw new InvalidOperationException(message.Trim());
            }
            if (!File.Exists(outputPath)) throw new IOException("The engine finished but the output workbook was not created.");
            return standardOutput;
        }
    }

    private static string FindNodeRuntime(string appDirectory)
    {
        string portableNode = Path.Combine(appDirectory, "runtime", "node.exe");
        if (File.Exists(portableNode)) return portableNode;

        string pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        string[] folders = pathValue.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string folder in folders)
        {
            try
            {
                string candidate = Path.Combine(folder.Trim(), "node.exe");
                if (File.Exists(candidate)) return candidate;
            }
            catch { }
        }
        throw new FileNotFoundException("Node.js runtime was not found on this computer.");
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
