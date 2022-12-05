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
    public partial class InsertClients : Form
    {
        private SqlConnection sqlConnection = null;
        public InsertClients(SqlConnection connection)
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

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Close();
        }

        private async void btnAdd_Click(object sender, EventArgs e)
        {
            SqlCommand insertClientsCommand = new SqlCommand("INSERT INTO [Clients] (Surname, Name, MiddleName, Budget, Area, Rooms, Floor, Sqaure, Email) VALUES(@Surname, @Name, @MiddleName, @Budget, @Area, @Rooms, @Floor, @Sqaure, @Email)", sqlConnection);
            insertClientsCommand.Parameters.AddWithValue("Surname", Convert.ToString(textBox1.Text));
            insertClientsCommand.Parameters.AddWithValue("Name", Convert.ToString(textBox2.Text));
            insertClientsCommand.Parameters.AddWithValue("Middlename", Convert.ToString(textBox3.Text));
            insertClientsCommand.Parameters.AddWithValue("Budget", Convert.ToDecimal(textBox4.Text));
            insertClientsCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox5.Text));
            insertClientsCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox6.Text));
            insertClientsCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox7.Text));
            insertClientsCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox8.Text));
            insertClientsCommand.Parameters.AddWithValue("Email", Convert.ToString(textBox9.Text));
            try
            {
                await insertClientsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных

                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            DataBank.SNMClients = textBox1.Text + " " + textBox2.Text + " " + textBox3.Text;
            DataBank.BudgetClients = textBox4.Text;
            DataBank.AreaClients = textBox5.Text;
            DataBank.RoomsClients = textBox6.Text;
            DataBank.FloorClients = textBox7.Text;
            DataBank.SqaureClients = textBox8.Text;
            DataBank.EmailClients = textBox9.Text;
            InsertTransactionsClients insertTransactions = new InsertTransactionsClients(sqlConnection);
            insertTransactions.Show();

            SqlCommand insertClientsCommand = new SqlCommand("INSERT INTO [Clients] (Surname, Name, MiddleName, Budget, Area, Rooms, Floor, Sqaure, Email) VALUES(@Surname, @Name, @MiddleName, @Budget, @Area, @Rooms, @Floor, @Sqaure, @Email)", sqlConnection);
            insertClientsCommand.Parameters.AddWithValue("Surname", Convert.ToString(textBox1.Text));
            insertClientsCommand.Parameters.AddWithValue("Name", Convert.ToString(textBox2.Text));
            insertClientsCommand.Parameters.AddWithValue("Middlename", Convert.ToString(textBox3.Text));
            insertClientsCommand.Parameters.AddWithValue("Budget", Convert.ToDecimal(textBox4.Text));
            insertClientsCommand.Parameters.AddWithValue("Area", Convert.ToString(textBox5.Text));
            insertClientsCommand.Parameters.AddWithValue("Rooms", Convert.ToInt32(textBox6.Text));
            insertClientsCommand.Parameters.AddWithValue("Floor", Convert.ToInt32(textBox7.Text));
            insertClientsCommand.Parameters.AddWithValue("Sqaure", Convert.ToDecimal(textBox8.Text));
            insertClientsCommand.Parameters.AddWithValue("Email", Convert.ToString(textBox9.Text));
            try
            {
                await insertClientsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}
