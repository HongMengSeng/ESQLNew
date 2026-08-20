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
            private readonly HashSet<string> _usedFields;
            private bool _updating;

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
                _grid.Columns.Add(fieldCol);
                _grid.CellValueChanged += (s, e) =>
                {
                    if (e.ColumnIndex == 1 && !_updating)
                        RefreshUsedFields();
                };
                _grid.CurrentCellDirtyStateChanged += (s, e) =>
                {
                    if (_grid.IsCurrentCellDirty && _grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == 1)
                        _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                };

                var presets = ColumnMapStore.ResolvePresets(excelHeaders, currentMap, _fieldNames);
                _usedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var p in presets)
                    if (!string.IsNullOrEmpty(p.Value))
                        _usedFields.Add(p.Value);
                foreach (var p in presets)
                {
                    var cell = new DataGridViewComboBoxCell();
                    var items = ColumnMapStore.AvailableFields(_fieldNames, _usedFields, p.Value);
                    cell.Items.Clear();
                    foreach (var it in items) cell.Items.Add(it);
                    cell.Value = p.Value;
                    _grid.Rows.Add(p.Key);
                    _grid.Rows[_grid.Rows.Count - 1].Cells[1] = cell;
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

            private void RefreshUsedFields()
            {
                _updating = true;
                try
                {
                    _usedFields.Clear();
                    for (int i = 0; i < _grid.Rows.Count; i++)
                    {
                        var val = _grid.Rows[i].Cells[1].Value as string;
                        if (!string.IsNullOrEmpty(val))
                            _usedFields.Add(val);
                    }
                    for (int i = 0; i < _grid.Rows.Count; i++)
                    {
                        var cell = _grid.Rows[i].Cells[1] as DataGridViewComboBoxCell;
                        if (cell == null) continue;
                        string current = _grid.Rows[i].Cells[1].Value as string;
                        var items = ColumnMapStore.AvailableFields(_fieldNames, _usedFields, current);
                        cell.Items.Clear();
                        foreach (var it in items) cell.Items.Add(it);
                        if (cell.Value == null || !items.Contains(cell.Value as string))
                            cell.Value = current;
                    }
                }
                finally
                {
                    _updating = false;
                }
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
