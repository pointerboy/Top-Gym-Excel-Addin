using System;
using System.Windows.Forms;

namespace Top_Gym_Excel
{
    public partial class CustomTaskPaneControl : UserControl
    {
        public ListBox MarkedUsersListBox;

        private void CustomTaskPaneControl_Load(object sender, EventArgs e) { }

        public CustomTaskPaneControl()
        {
            InitializeComponent();
            MarkedUsersListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Arial", 10)
            };
            this.Controls.Add(MarkedUsersListBox);
        }
    }
}
