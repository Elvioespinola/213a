using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace RealEstateAgency.Forms_of_action
{
    public partial class UpdateClients : Form
    {
        private SqlConnection sqlConnection = null;
        private int id;
        public UpdateClients(SqlConnection connection, int id)
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

        private async void UpdateClients_Load(object sender, EventArgs e)
            {
            SqlCommand getClientsInfoCommand = new SqlCommand("SELECT [Surname], [Name], [MiddleName], [Budget], [Area], [Rooms], [Floor], [Sqaure], [Email] FROM [Clients] WHERE [Id]=@Id", sqlConnection);
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
                    textBox4.Text = Convert.ToString(sqlReader["Budget"]);
                    textBox5.Text = Convert.ToString(sqlReader["Area"]);
                    textBox6.Text = Convert.ToString(sqlReader["Rooms"]);
                    textBox7.Text = Convert.ToString(sqlReader["Floor"]);
                    textBox8.Text = Convert.ToString(sqlReader["Sqaure"]);
                    textBox9.Text = Convert.ToString(sqlReader["Email"]);
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
            SqlCommand updateClientsInfoCommand = new SqlCommand("UPDATE [Clients] SET [Surname]=@Surname, [Name]=@Name, [MiddleName]=@MiddleName, [Budget]=@Budget, [Area]=@Area, [Rooms]=@Rooms, [Floor]=@Floor, [Sqaure]=@Sqaure, [Email]=@Email WHERE [Id]=@Id", sqlConnection);
            updateClientsInfoCommand.Parameters.AddWithValue("Id", id);
            updateClientsInfoCommand.Parameters.AddWithValue("Surname", Convert.ToString(textBox1.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Name", Convert.ToString(textBox2.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("MiddleName", Convert.ToString(textBox3.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Budget", Convert.ToDecimal(textBox4.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox5.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox6.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox7.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox8.Text));
            updateClientsInfoCommand.Parameters.AddWithValue("Email", Convert.ToString(textBox9.Text));
            try
            {
                await updateClientsInfoCommand.ExecuteNonQueryAsync();

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
