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
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        public void FlashWarning(string message)
        {
            listBox2.Items.Add(message);
            // Start the flashing effect
            Timer timer = new Timer();
            timer.Interval = 500; // Flash every 500ms
            timer.Tick += (sender, e) =>
            {
                // Toggle the background color between red and transparent
                this.BackColor = this.BackColor == Color.Transparent ? Color.Red : Color.Transparent;
            };

            timer.Start();

            // Stop flashing after 5 seconds and remove the label
            Task.Delay(30000).ContinueWith(t =>
            {
                this.Invoke(new Action(() =>
                {
                    timer.Stop();
                    this.BackColor = Color.Transparent; // Reset background
                }));
            });
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
        public void UpdateArrivalList(List<string> updatedArrivalList)
        {
            if (listBox1 != null)
            {
                listBox1.Items.Clear();
                foreach (var user in updatedArrivalList)
                {
                    listBox1.Items.Add(user);
                }
            }
            else
            {
                MessageBox.Show("The ListBox is not initialized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
