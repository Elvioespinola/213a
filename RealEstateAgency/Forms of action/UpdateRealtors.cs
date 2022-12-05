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
    public partial class UpdateRealtors : Form
    {
        private SqlConnection sqlConnection = null;
        private int id;
        public UpdateRealtors(SqlConnection connection, int id)
        {
            InitializeComponent();
            sqlConnection = connection;
            this.id = id;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private async void UpdateRealtors_Load(object sender, EventArgs e)
        {
            SqlCommand getClientsInfoCommand = new SqlCommand("SELECT [Surname], [Name], [MiddleName], [PercentFromDeal], [Email] FROM [Realtors] WHERE [Id]=@Id", sqlConnection);
            getClientsInfoCommand.Parameters.AddWithValue("Id", id);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await getClientsInfoCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = Convert.ToString(sqlReader["Surname"]);
                    textBox2.Text = Convert.ToString(sqlReader["Name"]);
                    textBox3.Text = Convert.ToString(sqlReader["MiddleName"]);
                    textBox4.Text = Convert.ToString(sqlReader["PercentFromDeal"]);
                    textBox5.Text = Convert.ToString(sqlReader["Email"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null && !sqlReader.IsClosed)
                    sqlReader.Close();
            }
        }

        private async void btnAdd_Click(object sender, EventArgs e)
        {
            SqlCommand updateRealtorsInfoCommand = new SqlCommand("UPDATE [Realtors] SET [Surname]=@Surname, [Name]=@Name, [MiddleName]=@MiddleName, [PercentFromDeal]=@PercentFromDeal, [Email]=@Email WHERE [Id]=@Id", sqlConnection);
            updateRealtorsInfoCommand.Parameters.AddWithValue("Id", id);
            updateRealtorsInfoCommand.Parameters.AddWithValue("Surname", Convert.ToString(textBox1.Text));
            updateRealtorsInfoCommand.Parameters.AddWithValue("Name", Convert.ToString(textBox2.Text));
            updateRealtorsInfoCommand.Parameters.AddWithValue("MiddleName", Convert.ToString(textBox3.Text));
            updateRealtorsInfoCommand.Parameters.AddWithValue("PercentFromDeal", Convert.ToInt32(textBox4.Text));
            updateRealtorsInfoCommand.Parameters.AddWithValue("Email", Convert.ToString(textBox5.Text));
            try
            {
                await updateRealtorsInfoCommand.ExecuteNonQueryAsync();

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
