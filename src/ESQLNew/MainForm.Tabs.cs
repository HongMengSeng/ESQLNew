using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ESQLNew.Core;
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
            page.Controls.Add(new Label
            {
                Text = "待实现",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            });
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
    }
}