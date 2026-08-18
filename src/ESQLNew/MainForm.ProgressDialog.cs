using System.Drawing;
using System.Windows.Forms;

namespace ESQLNew
{
    public partial class MainForm
    {
        private class ProgressDialog : Form
        {
            public ProgressBar Bar { get; private set; }
            public Label StepLabel { get; private set; }

            public ProgressDialog(string title)
            {
                Text = title;
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                ControlBox = false;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(360, 100);

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(10) };
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

                Bar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 4, Value = 0 };
                StepLabel = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Text = "正在准备…" };

                layout.Controls.Add(Bar, 0, 0);
                layout.Controls.Add(StepLabel, 0, 1);
                Controls.Add(layout);

                ShowInTaskbar = false;
            }

            public void ShowStep(int step, string text)
            {
                if (step < Bar.Minimum) step = Bar.Minimum;
                if (step > Bar.Maximum) step = Bar.Maximum;
                Bar.Value = step;
                StepLabel.Text = text;
            }
        }
    }
}
