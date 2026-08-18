using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ESQLNew.Core;
using ESQLNew.Import;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class ColumnMapDialog : Form
        {
            private readonly DataGridView _grid;
            private readonly List<string> _fieldNames;

            public Dictionary<string, string> Result { get; private set; }

            public ColumnMapDialog(string table, IList<string> excelHeaders,
                IList<ColumnInfo> tableColumns, Dictionary<string, string> currentMap)
            {
                Text = "编辑列映射 - " + table;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(480, 420);

                _fieldNames = new List<string>();
                foreach (var c in tableColumns)
                    _fieldNames.Add(c.Name);

                _grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    RowHeadersVisible = false,
                    BackgroundColor = SystemColors.Window,
                    EditMode = DataGridViewEditMode.EditOnEnter
                };

                _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Excel 表头", ReadOnly = true });
                var fieldCol = new DataGridViewComboBoxColumn { HeaderText = "英文字段", DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton, FlatStyle = FlatStyle.Flat };
                _grid.EditingControlShowing += (s, e) =>
                {
                    var combo = e.Control as DataGridViewComboBoxEditingControl;
                    if (combo != null)
                        combo.DropDownStyle = ComboBoxStyle.DropDownList;
                };
                foreach (var f in _fieldNames)
                    fieldCol.Items.Add(f);
                _grid.Columns.Add(fieldCol);

                for (int i = 0; i < excelHeaders.Count; i++)
                {
                    string header = excelHeaders[i] == null ? "" : excelHeaders[i].Trim();
                    if (header.Length == 0) continue;
                    string preset = null;
                    if (currentMap != null)
                        currentMap.TryGetValue(header, out preset);
                    if (preset != null && !_fieldNames.Contains(preset))
                        preset = null;
                    _grid.Rows.Add(header, preset);
                }

                var save = new Button { Text = "保存", Width = 90, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Width = 90, DialogResult = DialogResult.Cancel };
                var flow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.RightToLeft,
                    WrapContents = false,
                    Height = 40
                };
                flow.Controls.Add(save);
                flow.Controls.Add(cancel);
                flow.BringToFront();

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
                layout.Controls.Add(_grid, 0, 0);
                layout.Controls.Add(flow, 0, 1);
                Controls.Add(layout);

                AcceptButton = save;
                CancelButton = cancel;
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (DialogResult == DialogResult.OK)
                    Result = CollectResult();
                base.OnFormClosing(e);
            }

            private Dictionary<string, string> CollectResult()
            {
                var headers = new List<string>();
                var fields = new List<string>();
                for (int i = 0; i < _grid.Rows.Count; i++)
                {
                    var header = _grid.Rows[i].Cells[0].Value as string;
                    var field = _grid.Rows[i].Cells[1].Value as string;
                    headers.Add(header);
                    fields.Add(field);
                }
                return ColumnMapStore.BuildMap(headers, fields);
            }
        }
    }
}
