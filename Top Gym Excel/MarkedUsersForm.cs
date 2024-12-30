using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Top_Gym_Excel
{
    public partial class MarkedUsersForm : Form
    {
        private void MarkedUsersForm_Load(object sender, EventArgs e)
        {

        }

        public List<string> MarkedUsers { get; set; } = new List<string>();

        public MarkedUsersForm()
        {
            InitializeComponent();
            this.Text = "Marked Users for Today";
            this.Size = new System.Drawing.Size(400, 300);
            this.StartPosition = FormStartPosition.CenterScreen;

            // ListBox to show marked users
            ListBox markedUsersListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Arial", 10),
                ForeColor = System.Drawing.Color.Black
            };
            this.Controls.Add(markedUsersListBox);
        }

        public void UpdateList(List<string> markedUsers)
        {
            // Find the ListBox control added to the form
            ListBox markedUsersListBox = this.Controls[0] as ListBox;
            if (markedUsersListBox != null)
            {
                MarkedUsers.Clear();
                MarkedUsers.AddRange(markedUsers);
                markedUsersListBox.DataSource = null;  // Clear existing data source
                markedUsersListBox.DataSource = MarkedUsers;  // Set new data source
            }
        }
    }
}
