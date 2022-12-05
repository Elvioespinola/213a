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
    public partial class UpdateApartments : Form
    {
        private SqlConnection sqlConnection = null;
        private int id;
        public UpdateApartments(SqlConnection connection, int id)
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

        private async void UpdateApartments_Load(object sender, EventArgs e)
        {
            SqlCommand getApartmentsInfoCommand = new SqlCommand("SELECT [Address], [Value], [Area], [Rooms], [Floor], [Sqaure] FROM [Apartments] WHERE [Id]=@Id", sqlConnection);
            getApartmentsInfoCommand.Parameters.AddWithValue("Id", id);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await getApartmentsInfoCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = Convert.ToString(sqlReader["Address"]);
                    textBox2.Text = Convert.ToString(sqlReader["Value"]);
                    textBox3.Text = Convert.ToString(sqlReader["Area"]);
                    textBox4.Text = Convert.ToString(sqlReader["Rooms"]);
                    textBox5.Text = Convert.ToString(sqlReader["Floor"]);
                    textBox6.Text = Convert.ToString(sqlReader["Sqaure"]);
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
            SqlCommand updateApartmentsInfoCommand = new SqlCommand("UPDATE [Apartments] SET [Address]=@Address, [Value]=@Value, [Area]=@Area, [Rooms]=@Rooms, [Floor]=@Floor, [Sqaure]=@Sqaure WHERE [Id]=@Id", sqlConnection);
            updateApartmentsInfoCommand.Parameters.AddWithValue("Id", id);
            updateApartmentsInfoCommand.Parameters.AddWithValue("Address", Convert.ToString(textBox1.Text));
            updateApartmentsInfoCommand.Parameters.AddWithValue("Value", Convert.ToDecimal(textBox2.Text));
            updateApartmentsInfoCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox3.Text));
            updateApartmentsInfoCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox4.Text));
            updateApartmentsInfoCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox5.Text));
            updateApartmentsInfoCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox6.Text));
            try
            {
                await updateApartmentsInfoCommand.ExecuteNonQueryAsync();

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
