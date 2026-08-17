using System.Drawing;
using System.Windows.Forms;

namespace ESQLNew
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            Text = "ESQLNew - Excel 批量导入";
            ClientSize = new Size(560, 480);
            MinimumSize = new Size(500, 400);
            BuildTabs();
        }
    }
}