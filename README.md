
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Collections.Generic;

namespace EC_Scan_Report
{
    public class MainForm : Form
    {
        private Button generateButton;
        private CheckBox openAfterCheckBox;
        private ProgressBar progressBar;
        private Button cancelButton;
        private Label statusLabel;
        private CancellationTokenSource? cts;
        private Label yearsLabel;
        private TextBox yearsTextBox; // e.g., 2024,2025,2026

        public MainForm()
        {
            Text = "EC Scan Report Generator";
            Width = 620;
            Height = 260;
            var panel = new FlowLayoutPanel{ Dock = DockStyle.Fill, Padding = new Padding(12), FlowDirection = FlowDirection.TopDown, AutoSize = true };
            openAfterCheckBox = new CheckBox { Text = "Open report after generation", AutoSize = true };
            yearsLabel = new Label { Text = "Years (comma-separated, optional):", AutoSize = true };
            yearsTextBox = new TextBox { Width = 560 };
            progressBar = new ProgressBar { Width = 560, Height = 22, Minimum = 0, Maximum = 100 };
            statusLabel = new Label { AutoSize = true, Text = "Idle" };
            generateButton = new Button { Text = "Generate Report", Width = 240, Height = 34 };
            cancelButton = new Button { Text = "Cancel", Width = 120, Height = 34, Enabled = false };

            generateButton.Click += GenerateButton_Click;
            cancelButton.Click += (s, e) => cts?.Cancel();

            panel.Controls.Add(openAfterCheckBox);
            panel.Controls.Add(yearsLabel);
            panel.Controls.Add(yearsTextBox);
            panel.Controls.Add(generateButton);
            panel.Controls.Add(cancelButton);
            panel.Controls.Add(progressBar);
            panel.Controls.Add(statusLabel);
            Controls.Add(panel);
        }

        private async void GenerateButton_Click(object? sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select the top-level ECs folder (e.g., ECs (Engineering Changes))";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;
                    var years = ParseYears(yearsTextBox.Text);
                    cts = new CancellationTokenSource();
                    cancelButton.Enabled = true;
                    generateButton.Enabled = false;
                    try
                    {
                        var resultPath = await Task.Run(() => GenerateReport(folderPath, years, cts.Token));
                        if (!string.IsNullOrEmpty(resultPath))
                        {
                            MessageBox.Show($"Report generated: {resultPath}", "Success");
                            if (openAfterCheckBox.Checked)
                            {
                                Process.Start(new ProcessStartInfo { FileName = resultPath, UseShellExecute = true });
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        MessageBox.Show("Scan cancelled.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message, "Failure");
                    }
                    finally
                    {
                        cancelButton.Enabled = false;
                        generateButton.Enabled = true;
                        statusLabel.Text = "Idle";
                        progressBar.Value = 0;
                        cts = null;
                    }
                }
            }
        }

        private List<string> ParseYears(string input)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(input)) return list; // empty means all years
            var parts = input.Split(new[]{',',';',' '}, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var s = p.Trim();
                if (s.Length == 4 && s.All(char.IsDigit)) list.Add(s);
            }
            return list;
        }

        private string GenerateReport(string folderPath, List<string> years, CancellationToken token)
        {
            var allFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories)
                .Where(p => !Path.GetFileName(p).StartsWith("~$") && !p.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase));

            // Apply year filter: include only paths that contain \{year}\ in their folder structure
            var files = (years == null || years.Count == 0)
                ? allFiles.ToArray()
                : allFiles.Where(p => years.Any(y => p.IndexOf(Path.DirectorySeparatorChar + y + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) >= 0)).ToArray();

            int total = files.Length;
            if (total == 0) return string.Empty;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("EC Summary");
            var headers = new [] { "Source File", "EC#", "Originator", "Origination Date", "Priority", "Change Description", "ENG", "MFG", "QUALITY", "PURCHASING", "SALES", "MARKETING" };
            for (int i = 0; i < headers.Length; i++) ws.Cell(1, i+1).Value = headers[i];

            int row = 2;
            int processed = 0;

            foreach (var file in files)
            {
                token.ThrowIfCancellationRequested();
                var extracted = ExtractFromFirstSheet(file);
                ws.Cell(row, 1).Value = Path.GetFileName(file);
                ws.Cell(row, 2).Value = extracted.GetValueOrDefault("EC#", "");
                ws.Cell(row, 3).Value = extracted.GetValueOrDefault("Originator", "");
                ws.Cell(row, 4).Value = extracted.GetValueOrDefault("Origination Date", "");
                ws.Cell(row, 5).Value = extracted.GetValueOrDefault("Priority", "");
                ws.Cell(row, 6).Value = extracted.GetValueOrDefault("Change Description", "");
                ws.Cell(row, 7).Value = extracted.GetValueOrDefault("ENG", "");
                ws.Cell(row, 8).Value = extracted.GetValueOrDefault("MFG", "");
                ws.Cell(row, 9).Value = extracted.GetValueOrDefault("QUALITY", "");
                ws.Cell(row, 10).Value = extracted.GetValueOrDefault("PURCHASING", "");
                ws.Cell(row, 11).Value = extracted.GetValueOrDefault("SALES", "");
                ws.Cell(row, 12).Value = extracted.GetValueOrDefault("MARKETING", "");
                row++;

                processed++;
                int percent = (int)Math.Round((processed * 100.0) / total);
                UpdateProgress(percent, $"Processing {processed}/{total}: {Path.GetFileName(file)}");
            }

            ws.Columns().AdjustToContents();

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string suffix = (years == null || years.Count == 0) ? "ALL" : string.Join("-", years);
            string reportPath = Path.Combine(folderPath, $"Consolidated_EC_Summary_{suffix}_{timestamp}.xlsx");
            wb.SaveAs(reportPath);
            return reportPath;
        }

        private Dictionary<string, string> ExtractFromFirstSheet(string filePath)
        {
            var dict = new Dictionary<string, string>
            {
                ["EC#"] = "",
                ["Originator"] = "",
                ["Origination Date"] = "",
                ["Priority"] = "",
                ["Change Description"] = "",
                ["ENG"] = "",
                ["MFG"] = "",
                ["QUALITY"] = "",
                ["PURCHASING"] = "",
                ["SALES"] = "",
                ["MARKETING"] = ""
            };

            try
            {
                using (var fs = File.OpenRead(filePath))
                using (var wb = new ClosedXML.Excel.XLWorkbook(fs))
                {
                    var ws = wb.Worksheets.First();
                    int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                    for (int r = 1; r <= lastRow; r++)
                    {
                        var c1 = ws.Cell(r, 1).GetString().Trim();
                        var c2 = ws.Cell(r, 2).GetString().Trim();
                        var c3 = ws.Cell(r, 3).GetString().Trim();
                        var c4 = ws.Cell(r, 4).GetString().Trim();

                        if (c1 == "EC#") dict["EC#"] = c2;
                        else if (c1 == "Originator") dict["Originator"] = c2;
                        else if (c1 == "Origination Date") dict["Origination Date"] = c2;
                        else if (c1 == "Change Description") dict["Change Description"] = c2;

                        if (c1.Contains("Priority") || c2.Contains("Priority") || c3.Contains("Priority") || c4.Contains("Priority"))
                        {
                            if (new[]{c1,c2,c3,c4}.Any(v => v.Equals("URGENT", StringComparison.OrdinalIgnoreCase))) dict["Priority"] = "URGENT";
                            else if (new[]{c1,c2,c3,c4}.Any(v => v.Equals("ROUTINE", StringComparison.OrdinalIgnoreCase))) dict["Priority"] = "ROUTINE";
                        }

                        if (new[]{"ENG","MFG","QUALITY","PURCHASING","SALES","MARKETING"}.Contains(c1))
                        {
                            var initials = string.Join(" ", new[]{c2,c3,c4}.Where(s => !string.IsNullOrWhiteSpace(s)));
                            dict[c1] = initials.Trim();
                        }
                    }
                }
            }
            catch { }

            return dict;
        }

        private void UpdateProgress(int percent, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(percent, status)));
                return;
            }
            progressBar.Value = Math.Max(progressBar.Minimum, Math.Min(progressBar.Maximum, percent));
            statusLabel.Text = status;
        }
    }
}
