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
using ESQLNew.Logging;
using MySql.Data.MySqlClient;

namespace ESQLNew
{
    public partial class MainForm
    {
        private TextBox _serverTextBox;
        private TextBox _portTextBox;
        private TextBox _userTextBox;
        private TextBox _passwordTextBox;
        private ComboBox _databaseComboBox;
        private Button _dbRefreshButton;
        private CheckBox _showSysDbCheckBox;
        private ComboBox _tableComboBox;
        private Button _tableRefreshButton;
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
        private Button _previewButton;
        private Button _editMapButton;
        private Button _sheetSelectButton;
        private Button _importButton;
        private IList<string> _selectedSheets;
        private string _dedupKey;
        private string _previewSheet;
        private Label _statusLabel;
        private Label _countLabel;
        private ProgressBar _progressBar;
        private DataGridView _mappingGrid;
        private DataGridView _sampleGrid;

        private TabControl _tabs;
        private TabPage _logPage;
        private DataGridView _logGrid;
        private DataGridView _failGrid;
        private Button _logRefreshButton;
        private Button _logPrevButton;
        private Button _logNextButton;
        private Label _logPageLabel;
        private int _logPageNumber;
        private SqLiteLogStore _logStore;
        private bool _loadingConfig;

        private SqLiteLogStore LogStore
        {
            get
            {
                if (_logStore == null)
                {
                    _logStore = new SqLiteLogStore(Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "ESQLNew",
                        "logs.db"));
                }
                return _logStore;
            }
        }

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
            _tabs = new TabControl { Dock = DockStyle.Fill };
            _tabs.TabPages.Add(BuildConnectionTab());
            _tabs.TabPages.Add(BuildImportTab());
            _logPage = BuildLogTab();
            _tabs.TabPages.Add(_logPage);
            _tabs.SelectedIndexChanged += (s, e) =>
            {
                if (_tabs.SelectedTab == _logPage) RefreshLogList();
            };
            Controls.Add(_tabs);
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
            _dbRefreshButton.Click += (s, e) => ReloadDatabases();
            _showSysDbCheckBox.CheckedChanged += (s, e) =>
            {
                if (_loadingConfig) return;
                try
                {
                    var cfg = AppConfig.Load(ConfigPath);
                    cfg.ShowSystemDatabases = _showSysDbCheckBox.Checked;
                    cfg.Save(ConfigPath);
                }
                catch
                {
                }
                ReloadDatabases();
            };
            _databaseComboBox.SelectedIndexChanged += (s, e) =>
            {
                if (_databaseComboBox.SelectedItem != null)
                    ReloadTables(_databaseComboBox.SelectedItem.ToString());
            };
            _keepLogsCheckBox.CheckedChanged += (s, e) =>
            {
                if (_loadingConfig) return;
                try
                {
                    var cfg = AppConfig.Load(ConfigPath);
                    cfg.KeepLogs = _keepLogsCheckBox.Checked;
                    cfg.Save(ConfigPath);
                }
                catch
                {
                }
            };

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
            _tableComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend
            };
            _tableRefreshButton = new Button { Text = "刷新表", Dock = DockStyle.Fill };
            _previewButton = new Button { Text = "匹配预览", Dock = DockStyle.Fill };
            _editMapButton = new Button { Text = "编辑映射", Dock = DockStyle.Fill };
            _sheetSelectButton = new Button { Text = "选择工作表", Dock = DockStyle.Fill };
            _importButton = new Button { Text = "开始导入", Dock = DockStyle.Fill };
            _statusLabel = MakeLabel("");
            _countLabel = MakeLabel("已处理 0 / 总 0 / 成功 0 / 失败 0");
            _progressBar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100, Value = 0 };

            var tableFlow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, Margin = new Padding(0) };
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tableFlow.Controls.Add(_tableComboBox, 0, 0);
            tableFlow.Controls.Add(_previewButton, 1, 0);
            tableFlow.Controls.Add(_tableRefreshButton, 2, 0);
            tableFlow.Controls.Add(_editMapButton, 3, 0);

            var previewTabs = new TabControl { Dock = DockStyle.Fill };
            var mappingPage = new TabPage("字段映射");
            _mappingGrid = BuildPreviewGrid();
            mappingPage.Controls.Add(_mappingGrid);
            var samplePage = new TabPage("样例数据(前100行)");
            _sampleGrid = BuildPreviewGrid();
            samplePage.Controls.Add(_sampleGrid);
            previewTabs.TabPages.Add(mappingPage);
            previewTabs.TabPages.Add(samplePage);

            var fileFlow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            fileFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fileFlow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            fileFlow.Controls.Add(_fileTextBox, 0, 0);
            fileFlow.Controls.Add(_sheetSelectButton, 1, 0);

            layout.Controls.Add(_chooseFileButton, 0, 0);
            layout.Controls.Add(fileFlow, 1, 0);
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
            _sheetSelectButton.Click += ChooseSheet_Click;
            _tableRefreshButton.Click += (s, e) =>
            {
                if (_databaseComboBox.SelectedItem != null)
                    ReloadTables(_databaseComboBox.SelectedItem.ToString());
                else if (!string.IsNullOrWhiteSpace(_databaseComboBox.Text))
                    ReloadTables(_databaseComboBox.Text.Trim());
                else
                    MessageBox.Show(this, "请先选择数据库", "刷新表列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
            _previewButton.Click += Preview_Click;
            _editMapButton.Click += EditMapping_Click;
            _importButton.Click += StartImport_Click;

            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildLogTab()
        {
            var page = new TabPage("日志");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 65));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 35));

            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            _logRefreshButton = new Button { Text = "刷新", Width = 70 };
            _logPrevButton = new Button { Text = "上一页", Width = 70 };
            _logNextButton = new Button { Text = "下一页", Width = 70 };
            _logPageLabel = new Label { Text = "第 1 页", AutoSize = true, Margin = new Padding(10, 8, 0, 0) };
            toolbar.Controls.Add(_logRefreshButton);
            toolbar.Controls.Add(_logPrevButton);
            toolbar.Controls.Add(_logNextButton);
            toolbar.Controls.Add(_logPageLabel);

            _logGrid = BuildPreviewGrid();
            _logGrid.Columns.Add("Id", "Id");
            _logGrid.Columns[0].Visible = false;
            _logGrid.Columns.Add("Timestamp", "时间");
            _logGrid.Columns.Add("FileName", "文件");
            _logGrid.Columns.Add("TableName", "表名");
            _logGrid.Columns.Add("Total", "总数");
            _logGrid.Columns.Add("Succeeded", "成功");
            _logGrid.Columns.Add("Failed", "失败");
            _logGrid.Columns.Add("ElapsedMs", "耗时(ms)");
            _logGrid.Columns.Add("RowsPerSecond", "行每秒");

            _failGrid = BuildPreviewGrid();
            _failGrid.Columns.Add("RowNumber", "行号");
            _failGrid.Columns.Add("Message", "错误信息");

            layout.Controls.Add(toolbar, 0, 0);
            layout.Controls.Add(_logGrid, 0, 1);
            layout.Controls.Add(_failGrid, 0, 2);

            _logRefreshButton.Click += (s, e) => RefreshLogList();
            _logPrevButton.Click += (s, e) =>
            {
                if (_logPageNumber > 0)
                {
                    _logPageNumber--;
                    RefreshLogList();
                }
            };
            _logNextButton.Click += (s, e) =>
            {
                _logPageNumber++;
                RefreshLogList();
            };
            _logGrid.SelectionChanged += (s, e) => ShowFailures();

            page.Controls.Add(layout);
            return page;
        }

        private void RefreshLogList()
        {
            try
            {
                var cfg = AppConfig.Load(ConfigPath);
                if (cfg.LogRetentionDays > 0) LogStore.PurgeOld(cfg.LogRetentionDays);
                var rows = LogStore.Query(_logPageNumber, 100);
                _logGrid.Rows.Clear();
                foreach (var row in rows)
                    _logGrid.Rows.Add(row);
                _failGrid.Rows.Clear();
                _logPrevButton.Enabled = _logPageNumber > 0;
                _logNextButton.Enabled = rows.Count >= 100;
                _logPageLabel.Text = "第 " + (_logPageNumber + 1) + " 页";
            }
            catch
            {
            }
        }

        private void ShowFailures()
        {
            _failGrid.Rows.Clear();
            if (_logGrid.SelectedRows.Count == 0) return;
            var id = _logGrid.SelectedRows[0].Cells[0].Value;
            if (id == null) return;
            try
            {
                foreach (var row in LogStore.GetFailures((long)id))
                    _failGrid.Rows.Add(row);
            }
            catch
            {
            }
        }

        private void WriteImportLog(string file, string table, ImportResult result)
        {
            try
            {
                if (!AppConfig.Load(ConfigPath).KeepLogs) return;
                LogStore.LogImport(
                    DateTime.Now,
                    file,
                    table,
                    result.Total,
                    result.Succeeded,
                    result.Failed,
                    (long)result.Elapsed.TotalMilliseconds,
                    result.RowsPerSecond,
                    result.Failures);
                var cfg = AppConfig.Load(ConfigPath);
                if (cfg.LogRetentionDays > 0) LogStore.PurgeOld(cfg.LogRetentionDays);
                if (_tabs.SelectedTab == _logPage) RefreshLogList();
            }
            catch
            {
                _statusLabel.Text = "日志写入失败(导入已完成)";
            }
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
            _databaseComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend
            };
            _dbRefreshButton = new Button { Text = "刷新库", Width = 64 };
            _showSysDbCheckBox = new CheckBox
            {
                Text = "显示系统库",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 0)
            };
            var dbPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Padding = new Padding(0) };
            dbPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            dbPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
            dbPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            dbPanel.Controls.Add(_databaseComboBox, 0, 0);
            dbPanel.Controls.Add(_dbRefreshButton, 1, 0);
            dbPanel.Controls.Add(_showSysDbCheckBox, 2, 0);

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
            layout.Controls.Add(dbPanel, 1, 4);
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
            _databaseComboBox.Text = cfg.Database ?? "";
            _loadingConfig = true;
            _showSysDbCheckBox.Checked = cfg.ShowSystemDatabases;
            _rawTextBox.Text = cfg.RawConnectionString ?? "";
            _formRadio.Checked = !cfg.UseRaw;
            _rawRadio.Checked = cfg.UseRaw;
            _batchTextBox.Text = cfg.BatchSize.ToString();
            _commitTextBox.Text = cfg.CommitEvery.ToString();
            _keepLogsCheckBox.Checked = cfg.KeepLogs;
            _loadingConfig = false;
            UpdateModeFields();
        }

        private void UpdateModeFields()
        {
            bool form = _formRadio.Checked;
            _serverTextBox.Enabled = form;
            _portTextBox.Enabled = form;
            _userTextBox.Enabled = form;
            _passwordTextBox.Enabled = form;
            _databaseComboBox.Enabled = form;
            _dbRefreshButton.Enabled = form;
            _showSysDbCheckBox.Enabled = form;
            _generateButton.Enabled = form;
            _rawTextBox.ReadOnly = form;
        }

        private void ReloadDatabases()
        {
            try
            {
                var connStr = CurrentConnectionString();
                var all = DbMetadata.GetDatabases(connStr);
                var list = new List<string>();
                foreach (var db in all)
                    if (_showSysDbCheckBox.Checked || !DbMetadata.IsSystemDatabase(db))
                        list.Add(db);
                string current = _databaseComboBox.Text.Trim();
                _databaseComboBox.Items.Clear();
                foreach (var db in list)
                    _databaseComboBox.Items.Add(db);
                bool found = false;
                foreach (var db in list)
                    if (string.Equals(db, current, StringComparison.OrdinalIgnoreCase))
                        found = true;
                if (!found)
                {
                    if (list.Count > 0)
                        _databaseComboBox.Text = "";
                    if (_databaseComboBox.Items.Count > 0)
                        _databaseComboBox.SelectedIndex = -1;
                }
                else
                {
                    _databaseComboBox.Text = current;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "刷新库列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ReloadTables(string database)
        {
            if (_tableComboBox == null) return;
            if (string.IsNullOrWhiteSpace(database)) return;
            try
            {
                var tables = DbMetadata.GetTables(CurrentConnectionString(), database.Trim());
                _tableComboBox.Items.Clear();
                foreach (var t in tables)
                    _tableComboBox.Items.Add(t);
                _tableComboBox.Text = "";
            }
            catch (Exception ex)
            {
                _tableComboBox.Items.Clear();
                _tableComboBox.Text = "";
                MessageBox.Show(this, ex.Message, "刷新表列表", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
                _databaseComboBox.Text.Trim());
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
                        ReloadDatabases();
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
                Database = _databaseComboBox.Text.Trim(),
                RawConnectionString = _rawTextBox.Text,
                UseRaw = _rawRadio.Checked,
                BatchSize = ParseInt(_batchTextBox.Text, 2000),
                CommitEvery = ParseInt(_commitTextBox.Text, 5000),
                KeepLogs = _keepLogsCheckBox.Checked,
                ShowSystemDatabases = _showSysDbCheckBox.Checked
            };
            cfg.Save(ConfigPath);
            MessageBox.Show(this, "配置已保存", "保存配置", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static int ParseInt(string text, int fallback)
        {
            int value;
            return int.TryParse(text, out value) ? value : fallback;
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
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
                _tableComboBox.Text = "";
                _selectedSheets = null;
                _dedupKey = "";
                _previewSheet = null;
                _statusLabel.Text = "";
                try
                {
                    var names = ExcelStreamReader.SheetNames(ofd.FileName);
                    _selectedSheets = names;
                }
                catch (Exception ex)
                {
                    _statusLabel.Text = "无法读取工作表:" + ex.Message;
                }
            }
        }

        private void ChooseSheet_Click(object sender, EventArgs e)
        {
            string path = _fileTextBox.Text.Trim();
            if (path.Length == 0)
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var names = ExcelStreamReader.SheetNames(path);
                using (var dlg = new SheetSelectDialog(names, path))
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    _selectedSheets = dlg.SelectedSheets;
                    _dedupKey = dlg.DedupKey;
                    _previewSheet = _selectedSheets != null && _selectedSheets.Count > 0 ? _selectedSheets[0] : null;
                    _statusLabel.Text = "已选 " + (_selectedSheets != null ? _selectedSheets.Count : 0) + " 个工作表"
                        + (string.IsNullOrEmpty(_dedupKey) ? "" : ",去重键:" + _dedupKey);
                }
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "无法读取工作表:" + ex.Message;
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
            if (_selectedSheets == null || _selectedSheets.Count == 0)
            {
                MessageBox.Show(this, "请先选择工作表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _previewSheet = _selectedSheets[0];
            string table = _tableComboBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!TableExists(table))
            {
                MessageBox.Show(this, "表名不存在或已更改,请重新选择", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _previewButton.Enabled = false;
            _importButton.Enabled = false;
            _chooseFileButton.Enabled = false;
            _sheetSelectButton.Enabled = false;
            _tableComboBox.Enabled = false;
            _tableRefreshButton.Enabled = false;
            var dlg = new ProgressDialog("匹配预览中");
            dlg.Show(this);

            string connStr = CurrentConnectionString();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    if (IsDisposed || !IsHandleCreated) return;
                    if (!dlg.IsDisposed && dlg.IsHandleCreated)
                        dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(1, "读取 Excel 表头")));
                    var headers = ExcelStreamReader.ReadHeaders(path, _previewSheet);

                    if (IsDisposed || !IsHandleCreated) return;
                    if (!dlg.IsDisposed && dlg.IsHandleCreated)
                        dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(2, "读取目标表结构")));
                    var cols = ImportEngine.GetTableColumns(connStr, table);

                    if (IsDisposed || !IsHandleCreated) return;
                    if (!dlg.IsDisposed && dlg.IsHandleCreated)
                        dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(3, "匹配列映射")));
                    var fieldMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
                    var mappings = ColumnMapper.Map(headers, cols, fieldMap);

                    if (IsDisposed || !IsHandleCreated) return;
                    if (!dlg.IsDisposed && dlg.IsHandleCreated)
                        dlg.BeginInvoke((System.Action)(() => dlg.ShowStep(4, "加载样例数据")));
                    if (IsDisposed || !IsHandleCreated) return;
                    BeginInvoke((System.Action)(() =>
                    {
                        try
                        {
                            ShowMapping(mappings);
                            ShowSample(mappings);
                            int matched = 0;
                            foreach (var m in mappings)
                                if (m.Matched) matched++;
                            _statusLabel.Text = string.Format("共 {0} 列,匹配 {1} 列,未匹配 {2} 列", mappings.Count, matched, mappings.Count - matched);
                            if (ColumnMapStore.LastLoadCorrupt)
                                _statusLabel.Text += "  列映射配置读取失败,已使用内置映射";
                        }
                        catch (Exception uiEx)
                        {
                            MessageBox.Show(this, uiEx.Message, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            dlg.Close();
                            _previewButton.Enabled = true;
                            _importButton.Enabled = true;
                            _chooseFileButton.Enabled = true;
                            _sheetSelectButton.Enabled = true;
                            _tableComboBox.Enabled = true;
                            _tableRefreshButton.Enabled = true;
                        }
                    }));
                }
                catch (Exception ex)
                {
                    if (IsDisposed || !IsHandleCreated) return;
                    BeginInvoke((System.Action)(() =>
                    {
                        dlg.Close();
                        _previewButton.Enabled = true;
                        _importButton.Enabled = true;
                        _chooseFileButton.Enabled = true;
                        _sheetSelectButton.Enabled = true;
                        _tableComboBox.Enabled = true;
                        _tableRefreshButton.Enabled = true;
                        MessageBox.Show(this, ex.Message, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void EditMapping_Click(object sender, EventArgs e)
        {
            string path = _fileTextBox.Text.Trim();
            if (path.Length == 0)
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_selectedSheets == null || _selectedSheets.Count == 0)
            {
                MessageBox.Show(this, "请先选择工作表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string table = _tableComboBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                var headers = ExcelStreamReader.ReadHeaders(path, _previewSheet);
                IList<ColumnInfo> cols = null;
                bool columnsLoaded = false;
                try
                {
                    cols = ImportEngine.GetTableColumns(CurrentConnectionString(), table);
                    columnsLoaded = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "无法读取目标表字段:" + ex.Message + "\r\n仍可编辑,但英文字段下拉为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                var currentMap = ColumnMapStore.GetMap(ColumnMapStore.ConfigPath, table);
                var dlg = new ColumnMapDialog(table, headers, cols ?? new List<ColumnInfo>(), currentMap);
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Result != null)
                {
                    if (!columnsLoaded && currentMap != null && currentMap.Count > 0)
                    {
                        var confirm = MessageBox.Show(this,
                            "未读取到目标表字段,保存将清空表 '" + table + "' 的现有映射。确定保存吗?",
                            "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (confirm != DialogResult.Yes) return;
                    }
                    var maps = ColumnMapStore.Load(ColumnMapStore.ConfigPath);
                    maps[table] = dlg.Result;
                    ColumnMapStore.Save(ColumnMapStore.ConfigPath, maps);
                    _statusLabel.Text = "列映射已保存,共 " + dlg.Result.Count + " 列映射";
                    Preview_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "编辑映射失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (_selectedSheets == null || _selectedSheets.Count == 0)
            {
                MessageBox.Show(this, "请先选择工作表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string table = _tableComboBox.Text.Trim();
            if (!IsValidTableName(table))
            {
                MessageBox.Show(this, "表名只能包含字母、数字和下划线", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!TableExists(table))
            {
                MessageBox.Show(this, "表名不存在或已更改,请重新选择", "表名无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string connStr = CurrentConnectionString();
            int batch = Clamp(ParseInt(_batchTextBox.Text, 2000), 500, 10000);
            int commit = Clamp(ParseInt(_commitTextBox.Text, 5000), 1000, 100000);

            _importButton.Enabled = false;
            _previewButton.Enabled = false;
            _chooseFileButton.Enabled = false;
            _sheetSelectButton.Enabled = false;
            _tableComboBox.Enabled = false;
            _tableRefreshButton.Enabled = false;
            _progressBar.Maximum = 100;
            _progressBar.Value = 0;
            _countLabel.Text = "已处理 0 / 总 0 / 成功 0 / 失败 0";

            var lastProgress = new ImportProgress();
            Action<ImportProgress> progressCb = p =>
            {
                lastProgress = p;
                if (IsDisposed || !IsHandleCreated) return;
                BeginInvoke((Action)(() => UpdateProgress(p)));
            };

            try
            {
                var result = await Task.Run(() =>
                    ImportEngine.RunMultiSheet(connStr, table, path, _selectedSheets, _dedupKey, batch, commit, progressCb, System.Threading.CancellationToken.None).Result);
                ShowReport(result);
                WriteImportLog(path, table, result);
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (lastProgress.Processed > 0)
                    msg += string.Format("\r\n\r\n已处理 {0} 行(成功 {1} / 失败 {2})",
                        lastProgress.Processed, lastProgress.Succeeded, lastProgress.Failed);
                MessageBox.Show(this, msg, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _importButton.Enabled = true;
                _previewButton.Enabled = true;
                _chooseFileButton.Enabled = true;
                _sheetSelectButton.Enabled = true;
                _tableComboBox.Enabled = true;
                _tableRefreshButton.Enabled = true;
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
            if (indices.Count == 0) return;
            int count = 0;
            foreach (var row in ExcelStreamReader.ReadRows(_fileTextBox.Text.Trim(), _previewSheet))
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

        private bool TableExists(string table)
        {
            if (_tableComboBox.Items.Count == 0) return true;
            foreach (var item in _tableComboBox.Items)
                if (string.Equals(item.ToString(), table, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
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
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
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