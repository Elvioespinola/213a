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
    public partial class InsertApartments : Form
    {
        private SqlConnection sqlConnection = null;
        public InsertApartments(SqlConnection connection)
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

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void panel1_MouseDown_1(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private async void btnAdd_Click_1(object sender, EventArgs e)
        {
            SqlCommand insertApartmentsCommand = new SqlCommand("INSERT INTO [Apartments] (Address, Value, Area, Rooms, Floor, Sqaure) VALUES(@Address, @Value, @Area, @Rooms, @Floor, @Sqaure)", sqlConnection);
            insertApartmentsCommand.Parameters.AddWithValue("Address", Convert.ToString(textBox1.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Value", Convert.ToDecimal(textBox2.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox3.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox4.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox5.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox6.Text));
            try
            {
                await insertApartmentsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            DataBank.AddressApartment = textBox1.Text;
            DataBank.ValueApartment = textBox2.Text;
            DataBank.AreaApartment = textBox3.Text;
            DataBank.RoomsApartment = textBox4.Text;
            DataBank.FloorApartment = textBox5.Text;
            DataBank.SqaureApartment = textBox6.Text;
            InsertTransactionsClients insertTransactions = new InsertTransactionsClients(sqlConnection);
            insertTransactions.Show();

            SqlCommand insertApartmentsCommand = new SqlCommand("INSERT INTO [Apartments] (Address, Value, Area, Rooms, Floor, Sqaure) VALUES(@Address, @Value, @Area, @Rooms, @Floor, @Sqaure)", sqlConnection);
            insertApartmentsCommand.Parameters.AddWithValue("Address", Convert.ToString(textBox1.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Value", Convert.ToDecimal(textBox2.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox3.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox4.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox5.Text));
            insertApartmentsCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox6.Text));
            try
            {
                await insertApartmentsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}