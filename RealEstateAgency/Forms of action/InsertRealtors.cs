using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace RealEstateAgency.Forms_of_action
{
    public partial class InsertRealtors : Form
    {
        private SqlConnection sqlConnection = null;
        public InsertRealtors(SqlConnection connection)
        {
            InitializeComponent();
            sqlConnection = connection;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private async void btnAdd_Click(object sender, EventArgs e)
        {
            SqlCommand insertRealtorsCommand = new SqlCommand("INSERT INTO [Realtors] (Surname, Name, MiddleName, PercentFromDeal, Email) VALUES(@Surname, @Name, @MiddleName, @PercentFromDeal , @Email)", sqlConnection);
            insertRealtorsCommand.Parameters.AddWithValue("Surname", Convert.ToString(textBox1.Text));
            insertRealtorsCommand.Parameters.AddWithValue("Name", Convert.ToString(textBox2.Text));
            insertRealtorsCommand.Parameters.AddWithValue("Middlename", Convert.ToString(textBox3.Text));
            insertRealtorsCommand.Parameters.AddWithValue("PercentFromDeal", Convert.ToInt32(textBox4.Text));
            insertRealtorsCommand.Parameters.AddWithValue("Email", Convert.ToString(textBox5.Text));
            try
            {
                await insertRealtorsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
