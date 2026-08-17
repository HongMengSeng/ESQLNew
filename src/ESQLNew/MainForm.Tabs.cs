using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ESQLNew.Core;
using ESQLNew.Excel;
using ESQLNew.Import;
using MySql.Data.MySqlClient;

namespace ESQLNew
{
    public partial class MainForm
    {
        private TextBox _serverTextBox;
        private TextBox _portTextBox;
        private TextBox _userTextBox;
        private TextBox _passwordTextBox;
        private TextBox _databaseTextBox;
        private RadioButton _formRadio;
        private RadioButton _rawRadio;
        private TextBox _rawTextBox;
        private Button _generateButton;
        private Button _testButton;
        private Button _saveButton;
        private TextBox _batchTextBox;
        private TextBox _commitTextBox;
        private CheckBox _keepLogsCheckBox;

        private Button _chooseFileButton;
        private TextBox _fileTextBox;
        private TextBox _tableTextBox;
        private Button _previewButton;
        private Button _importButton;
        private Label _statusLabel;
        private Label _countLabel;
        private ProgressBar _progressBar;
        private DataGridView _mappingGrid;
        private DataGridView _sampleGrid;

        private static string ConfigPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ESQLNew",
                    "config.json");
            }
        }

        private void BuildTabs()
        {
            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(BuildConnectionTab());
            tabs.TabPages.Add(BuildImportTab());
            tabs.TabPages.Add(BuildLogTab());
            Controls.Add(tabs);
            LoadConfig();
        }

        private TabPage BuildConnectionTab()
        {
            var page = new TabPage("连接");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(8) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 300));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(BuildConnectionGroup(), 0, 0);
            layout.Controls.Add(BuildAdvancedGroup(), 0, 1);
            page.Controls.Add(layout);

            _formRadio.CheckedChanged += (s, e) => UpdateModeFields();
            _rawRadio.CheckedChanged += (s, e) => UpdateModeFields();
            _generateButton.Click += (s, e) => GenerateConnectionString();
            _testButton.Click += (s, e) => TestConnection();
            _saveButton.Click += (s, e) => SaveConfig();

            return page;
        }

        private TabPage BuildImportTab()
        {
            var page = new TabPage("导入");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 6, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            _chooseFileButton = new Button { Text = "选择文件", Dock = DockStyle.Fill };
            _fileTextBox = new TextBox { Dock = DockStyle.Fill, ReadOnly = true };
            _tableTextBox = new TextBox { Dock = DockStyle.Fill };
            _previewButton = new Button { Text = "匹配预览", Dock = DockStyle.Fill };
            _importButton = new Button { Text = "开始导入", Dock = DockStyle.Fill };
            _statusLabel = MakeLabel("");
            _countLabel = MakeLabel("已处理 0 / 总 0 / 成功 0 / 失败 0");
            _progressBar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100, Value = 0 };

            var tableFlow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.Controls.Add(_tableTextBox, 0, 0);
            tableFlow.Controls.Add(_previewButton, 1, 0);

            var previewTabs = new TabControl { Dock = DockStyle.Fill };
            var mappingPage = new TabPage("字段映射");
            _mappingGrid = BuildPreviewGrid();
            mappingPage.Controls.Add(_mappingGrid);
            var samplePage = new TabPage("样例数据(前100行)");
            _sampleGrid = BuildPreviewGrid();
            samplePage.Controls.Add(_sampleGrid);
            previewTabs.TabPages.Add(mappingPage);
            previewTabs.TabPages.Add(samplePage);

            layout.Controls.Add(_chooseFileButton, 0, 0);
            layout.Controls.Add(_fileTextBox, 1, 0);
            layout.Controls.Add(MakeLabel("目标表"), 0, 1);
            layout.Controls.Add(tableFlow, 1, 1);
            layout.SetColumnSpan(_statusLabel, 2);
            layout.Controls.Add(_statusLabel, 0, 2);
            layout.SetColumnSpan(previewTabs, 2);
            layout.Controls.Add(previewTabs, 0, 3);
            layout.SetColumnSpan(_progressBar, 2);
            layout.Controls.Add(_progressBar, 0, 4);
            layout.Controls.Add(_importButton, 0, 5);
            layout.Controls.Add(_countLabel, 1, 5);

            _chooseFileButton.Click += ChooseFile_Click;
            _previewButton.Click += Preview_Click;
            _importButton.Click += StartImport_Click;

            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildLogTab()
        {
            var page = new TabPage("日志");
            page.Controls.Add(new Label
            {
                Text = "待实现",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            });
            return page;
        }

        private GroupBox BuildConnectionGroup()
        {
            var group = new GroupBox { Text = "连接设置", Dock = DockStyle.Fill };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowCount = 8;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            _serverTextBox = new TextBox { Dock = DockStyle.Fill };
            _portTextBox = new TextBox { Dock = DockStyle.Fill };
            _userTextBox = new TextBox { Dock = DockStyle.Fill };
            _passwordTextBox = new TextBox { Dock = DockStyle.Fill, PasswordChar = '●' };
            _databaseTextBox = new TextBox { Dock = DockStyle.Fill };

            _formRadio = new RadioButton { Text = "表单模式", AutoSize = true, Margin = new Padding(0, 4, 8, 0) };
            _rawRadio = new RadioButton { Text = "连接串模式", AutoSize = true, Margin = new Padding(0, 4, 8, 0) };
            _generateButton = new Button { Text = "生成连接串" };
            var modeFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            modeFlow.Controls.Add(_formRadio);
            modeFlow.Controls.Add(_rawRadio);
            modeFlow.Controls.Add(_generateButton);

            _rawTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };

            _testButton = new Button { Text = "测试连接" };
            _saveButton = new Button { Text = "保存配置" };
            var buttonFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            buttonFlow.Controls.Add(_testButton);
            buttonFlow.Controls.Add(_saveButton);

            layout.Controls.Add(MakeLabel("服务器"), 0, 0);
            layout.Controls.Add(_serverTextBox, 1, 0);
            layout.Controls.Add(MakeLabel("端口"), 0, 1);
            layout.Controls.Add(_portTextBox, 1, 1);
            layout.Controls.Add(MakeLabel("用户名"), 0, 2);
            layout.Controls.Add(_userTextBox, 1, 2);
            layout.Controls.Add(MakeLabel("密码"), 0, 3);
            layout.Controls.Add(_passwordTextBox, 1, 3);
            layout.Controls.Add(MakeLabel("数据库"), 0, 4);
            layout.Controls.Add(_databaseTextBox, 1, 4);
            layout.SetColumnSpan(modeFlow, 2);
            layout.Controls.Add(modeFlow, 0, 5);
            layout.Controls.Add(MakeLabel("连接串"), 0, 6);
            layout.Controls.Add(_rawTextBox, 1, 6);
            layout.SetColumnSpan(buttonFlow, 2);
            layout.Controls.Add(buttonFlow, 0, 7);

            group.Controls.Add(layout);
            return group;
        }

        private GroupBox BuildAdvancedGroup()
        {
            var group = new GroupBox { Text = "高级", Dock = DockStyle.Fill };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowCount = 3;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));

            _batchTextBox = new TextBox { Dock = DockStyle.Fill };
            _commitTextBox = new TextBox { Dock = DockStyle.Fill };
            _keepLogsCheckBox = new CheckBox
            {
                Text = "保留日志",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Checked = true
            };

            layout.Controls.Add(MakeLabel("批大小"), 0, 0);
            layout.Controls.Add(_batchTextBox, 1, 0);
            layout.Controls.Add(MakeLabel("提交频率"), 0, 1);
            layout.Controls.Add(_commitTextBox, 1, 1);
            layout.Controls.Add(MakeLabel("日志"), 0, 2);
            layout.Controls.Add(_keepLogsCheckBox, 1, 2);

            group.Controls.Add(layout);
            return group;
        }

        private static Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private void LoadConfig()
        {
            var cfg = AppConfig.Load(ConfigPath);
            _serverTextBox.Text = cfg.Server ?? "";
            _portTextBox.Text = cfg.Port.ToString();
            _userTextBox.Text = cfg.User ?? "";
            _passwordTextBox.Text = cfg.Password ?? "";
            _databaseTextBox.Text = cfg.Database ?? "";
            _rawTextBox.Text = cfg.RawConnectionString ?? "";
            _formRadio.Checked = !cfg.UseRaw;
            _rawRadio.Checked = cfg.UseRaw;
            _batchTextBox.Text = cfg.BatchSize.ToString();
            _commitTextBox.Text = cfg.CommitEvery.ToString();
            _keepLogsCheckBox.Checked = cfg.KeepLogs;
            UpdateModeFields();
        }

        private void UpdateModeFields()
        {
            bool form = _formRadio.Checked;
            _serverTextBox.Enabled = form;
            _portTextBox.Enabled = form;
            _userTextBox.Enabled = form;
            _passwordTextBox.Enabled = form;
            _databaseTextBox.Enabled = form;
            _generateButton.Enabled = form;
            _rawTextBox.ReadOnly = form;
        }

        private string CurrentConnectionString()
        {
            if (_rawRadio.Checked)
                return _rawTextBox.Text.Trim();
            return MySqlConnectionBuilder.Build(
                _serverTextBox.Text.Trim(),
                ParseInt(_portTextBox.Text, 3306),
                _userTextBox.Text.Trim(),
                _passwordTextBox.Text,
                _databaseTextBox.Text.Trim());
        }

        private void GenerateConnectionString()
        {
            _rawTextBox.Text = CurrentConnectionString();
        }

        private void TestConnection()
        {
            string cs = CurrentConnectionString();
            _testButton.Enabled = false;
            Task.Run(() =>
            {
                string error = null;
                try
                {
                    using (var conn = new MySqlConnection(cs))
                    {
                        conn.Open();
                    }
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                }
                BeginInvoke((Action)(() =>
                {
                    _testButton.Enabled = true;
                    if (error == null)
                    {
                        MessageBox.Show(this, "连接成功", "✅ 连接成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(this, error, "❌ 连接失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }));
            });
        }

        private void SaveConfig()
        {
            var cfg = new AppConfig
            {
                Server = _serverTextBox.Text.Trim(),
                Port = ParseInt(_portTextBox.Text, 3306),
                User = _userTextBox.Text.Trim(),
                Password = _passwordTextBox.Text,
                Database = _databaseTextBox.Text.Trim(),
                RawConnectionString = _rawTextBox.Text,
                UseRaw = _rawRadio.Checked,
                BatchSize = ParseInt(_batchTextBox.Text, 2000),
                CommitEvery = ParseInt(_commitTextBox.Text, 5000),
                KeepLogs = _keepLogsCheckBox.Checked
            };
            cfg.Save(ConfigPath);
            MessageBox.Show(this, "配置已保存", "保存配置", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static int ParseInt(string text, int fallback)
        {
            int value;
            return int.TryParse(text, out value) ? value : fallback;
        }

        private void ChooseFile_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "选择 Excel 文件",
                Filter = "Excel 文件|*.xlsx;*.xls"
            })
            {
                if (ofd.ShowDialog(this) != DialogResult.OK) return;
                _fileTextBox.Text = ofd.FileName;
                _tableTextBox.Text = "";
                _statusLabel.Text = "";
                try
                {
                    var names = ExcelStreamReader.SheetNames(ofd.FileName);
                    if (names.Count > 0)
                        _tableTextBox.Text = names[0];
                }
                catch (Exception ex)
                {
                    _statusLabel.Text = "无法读取工作表:" + ex.Message;
                }
            }
        }

        private void Preview_Click(object sender, EventArgs e)
        {
            string path = _fileTextBox.Text.Trim();
            if (path.Length == 0)
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string table = _tableTextBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var headers = ExcelStreamReader.ReadHeaders(path);
                var cols = ImportEngine.GetTableColumns(CurrentConnectionString(), table);
                var mappings = ColumnMapper.Map(headers, cols);
                ShowMapping(mappings);
                ShowSample(mappings);
                int matched = 0;
                foreach (var m in mappings)
                    if (m.Matched) matched++;
                _statusLabel.Text = string.Format("共 {0} 列,匹配 {1} 列,未匹配 {2} 列", mappings.Count, matched, mappings.Count - matched);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void StartImport_Click(object sender, EventArgs e)
        {
            string path = _fileTextBox.Text.Trim();
            if (path.Length == 0)
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string table = _tableTextBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string connStr = CurrentConnectionString();
            int batch = ParseInt(_batchTextBox.Text, 2000);
            int commit = ParseInt(_commitTextBox.Text, 5000);

            _importButton.Enabled = false;
            _previewButton.Enabled = false;
            _chooseFileButton.Enabled = false;
            _tableTextBox.Enabled = false;
            _progressBar.Maximum = 100;
            _progressBar.Value = 0;
            _countLabel.Text = "已处理 0 / 总 0 / 成功 0 / 失败 0";

            Action<ImportProgress> progressCb = p =>
                BeginInvoke((Action)(() => UpdateProgress(p)));

            try
            {
                var result = await Task.Run(() =>
                    ImportEngine.Run(connStr, table, path, batch, commit, progressCb, System.Threading.CancellationToken.None).Result);
                ShowReport(result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _importButton.Enabled = true;
                _previewButton.Enabled = true;
                _chooseFileButton.Enabled = true;
                _tableTextBox.Enabled = true;
            }
        }

        private void UpdateProgress(ImportProgress p)
        {
            long max = Math.Max(1L, Math.Max(p.Total, p.Processed));
            long val = Math.Min(p.Processed, max);
            if (max > int.MaxValue) max = int.MaxValue;
            if (val > int.MaxValue) val = int.MaxValue;
            _progressBar.Maximum = (int)max;
            _progressBar.Value = (int)val;
            _countLabel.Text = string.Format("已处理 {0} / 总 {1} / 成功 {2} / 失败 {3}", p.Processed, p.Total, p.Succeeded, p.Failed);
        }

        private void ShowMapping(IList<ColumnMapping> mappings)
        {
            _mappingGrid.Columns.Clear();
            _mappingGrid.Columns.Add("ExcelColumn", "Excel 列");
            _mappingGrid.Columns.Add("TableField", "表字段");
            _mappingGrid.Columns.Add("Status", "匹配状态");
            _mappingGrid.Rows.Clear();
            foreach (var m in mappings)
                _mappingGrid.Rows.Add(m.ExcelColumn, m.TableField ?? "", m.Matched ? "匹配" : "未匹配");
        }

        private void ShowSample(IList<ColumnMapping> mappings)
        {
            var indices = new List<int>();
            for (int i = 0; i < mappings.Count; i++)
                if (mappings[i].Matched) indices.Add(i);
            _sampleGrid.Columns.Clear();
            foreach (var idx in indices)
                _sampleGrid.Columns.Add(mappings[idx].ExcelColumn, mappings[idx].ExcelColumn);
            _sampleGrid.Rows.Clear();
            int count = 0;
            foreach (var row in ExcelStreamReader.ReadRows(_fileTextBox.Text.Trim()))
            {
                var values = new object[indices.Count];
                for (int i = 0; i < indices.Count; i++)
                    values[i] = row[indices[i]] == null ? "" : Convert.ToString(row[indices[i]]);
                _sampleGrid.Rows.Add(values);
                if (++count >= 100) break;
            }
        }

        private static DataGridView BuildPreviewGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.Window
            };
        }

        private static bool IsValidTableName(string name)
        {
            return !string.IsNullOrEmpty(name) && Regex.IsMatch(name, "^[A-Za-z0-9_]+$");
        }

        private void ShowReport(ImportResult result)
        {
            var dlg = new ImportReportDialog(result);
            dlg.ExportRequested += () =>
            {
                using (var sfd = new SaveFileDialog
                {
                    Title = "导出失败明细",
                    Filter = "CSV 文件|*.csv",
                    DefaultExt = "csv",
                    AddExtension = true,
                    FileName = "failures_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv"
                })
                {
                    if (sfd.ShowDialog(this) != DialogResult.OK) return;
                    try
                    {
                        ExportFailures(result.Failures, sfd.FileName);
                        MessageBox.Show(this, "已导出到:\r\n" + sfd.FileName, "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, ex.Message, "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
            dlg.ShowDialog(this);
        }

        private static void ExportFailures(IList<RowFailure> failures, string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine("RowNumber,Message");
            foreach (var f in failures)
                sb.AppendLine(f.RowNumber.ToString() + "," + CsvEscape(f.Message));
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private static string CsvEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            if (value.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        private class ImportReportDialog : Form
        {
            public event Action ExportRequested;

            public ImportReportDialog(ImportResult result)
            {
                Text = "导入报告";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(360, 300);

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12) };
                layout.ColumnCount = 1;
                layout.RowCount = 3;
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

                var info = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical,
                    Text = BuildReportText(result)
                };
                layout.Controls.Add(info, 0, 0);

                var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
                var close = new Button { Text = "关闭", Width = 90 };
                close.Click += (s, e) => Close();
                flow.Controls.Add(close);
                if (result.Failures.Count > 0)
                {
                    var export = new Button { Text = "导出失败明细 CSV", Width = 140 };
                    export.Click += (s, e) => { if (ExportRequested != null) ExportRequested(); };
                    flow.Controls.Add(export);
                }
                layout.Controls.Add(flow, 0, 2);
                Controls.Add(layout);
            }

            private static string BuildReportText(ImportResult r)
            {
                return string.Format(
                    "导入完成\r\n\r\n总数: {0}\r\n成功: {1}\r\n失败: {2}\r\n耗时: {3:F1} 秒\r\n行每秒: {4:F1}\r\n批次数: {5}",
                    r.Total, r.Succeeded, r.Failed, r.Elapsed.TotalSeconds, r.RowsPerSecond, r.BatchCount);
            }
        }
    }
}