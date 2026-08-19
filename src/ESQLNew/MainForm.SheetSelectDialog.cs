using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ESQLNew.Excel;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class SheetSelectDialog : Form
        {
            private readonly IList<string> _sheetNames;
            private readonly string _filePath;
            private readonly CheckedListBox _list;
            private readonly ComboBox _dedup;

            public IList<string> SelectedSheets { get; private set; }
            public string DedupKey { get; private set; }

            public SheetSelectDialog(IList<string> sheetNames, string filePath)
            {
                Text = "选择工作表";
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ShowInTaskbar = false;
                ClientSize = new Size(360, 300);

                _sheetNames = sheetNames;
                _filePath = filePath;

                _list = new CheckedListBox { Dock = DockStyle.Fill };
                foreach (var n in sheetNames)
                    _list.Items.Add(n, true);
                _list.ItemCheck += (s, e) =>
                {
                    this.BeginInvoke((Action)RebuildDedup);
                };

                _dedup = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                _dedup.Items.Add("不去重");
                if (sheetNames.Count > 0)
                    PopulateDedup(sheetNames[0]);
                _dedup.SelectedIndex = 0;

                var all = new Button { Text = "全选", Dock = DockStyle.Fill };
                var none = new Button { Text = "全不选", Dock = DockStyle.Fill };
                var ok = new Button { Text = "确定", Dock = DockStyle.Fill, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Dock = DockStyle.Fill, DialogResult = DialogResult.Cancel };

                all.Click += (s, e) =>
                {
                    for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, true);
                };
                none.Click += (s, e) =>
                {
                    for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, false);
                };

                var btnFlow = new FlowLayoutPanel { Dock = DockStyle.Fill };
                btnFlow.Controls.Add(all);
                btnFlow.Controls.Add(none);
                btnFlow.Controls.Add(ok);
                btnFlow.Controls.Add(cancel);

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
                layout.Controls.Add(_list, 0, 0);
                layout.Controls.Add(_dedup, 0, 1);
                layout.Controls.Add(btnFlow, 0, 2);
                Controls.Add(layout);

                AcceptButton = ok;
                CancelButton = cancel;

                ok.Click += Ok_Click;
            }

            private void Ok_Click(object sender, EventArgs e)
            {
                var selected = new List<string>();
                for (int i = 0; i < _list.Items.Count; i++)
                    if (_list.GetItemChecked(i))
                        selected.Add((string)_list.Items[i]);
                SelectedSheets = selected;
                DedupKey = _dedup.SelectedIndex <= 0 ? "" : (string)_dedup.SelectedItem;
            }

            private void PopulateDedup(string sheetName)
            {
                string current = _dedup.SelectedIndex > 0 ? (string)_dedup.SelectedItem : null;
                _dedup.Items.Clear();
                _dedup.Items.Add("不去重");
                try
                {
                    var headers = ExcelStreamReader.ReadHeaders(_filePath, sheetName);
                    if (headers != null)
                        foreach (var h in headers)
                            if (!string.IsNullOrEmpty(h))
                                _dedup.Items.Add(h);
                }
                catch
                {
                }
                _dedup.SelectedIndex = 0;
                if (current != null)
                    for (int i = 1; i < _dedup.Items.Count; i++)
                        if (string.Equals((string)_dedup.Items[i], current, StringComparison.Ordinal))
                        {
                            _dedup.SelectedIndex = i;
                            break;
                        }
            }

            private void RebuildDedup()
            {
                if (IsDisposed) return;
                string first = null;
                for (int i = 0; i < _list.Items.Count; i++)
                    if (_list.GetItemChecked(i))
                    {
                        first = (string)_list.Items[i];
                        break;
                    }
                if (first != null)
                    PopulateDedup(first);
            }
        }
    }
}
