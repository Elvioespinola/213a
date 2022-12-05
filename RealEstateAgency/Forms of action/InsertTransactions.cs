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
    public partial class InsertTransactionsClients : Form
    {
        //переменные
        private SqlConnection sqlConnection = null;
        private int id;
        public InsertTransactionsClients(SqlConnection connection, int id)
        {
            InitializeComponent();
            sqlConnection = connection;
            this.id = id;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }
        
        public InsertTransactionsClients(SqlConnection connection)
        {
            sqlConnection = connection;
            InitializeComponent();
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
        }

        //интерфейс
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        //форма закрытия
        private void button2_Click(object sender, EventArgs e)
        {
            DataBank.AddressApartment = null;
            DataBank.ValueApartment = null;
            DataBank.AreaApartment = null;
            DataBank.RoomsApartment = null;
            DataBank.FloorApartment = null;
            DataBank.SqaureApartment = null;

            DataBank.SNMClients = null;
            DataBank.BudgetClients = null;
            DataBank.AreaClients = null;
            DataBank.RoomsClients = null;
            DataBank.FloorClients = null;
            DataBank.SqaureClients = null;
            DataBank.EmailClients = null;
            Close();
        }
        //захват окна
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        //закрытие
        private void btnDelete_Click(object sender, EventArgs e)
        {
            DataBank.AddressApartment = null;
            DataBank.ValueApartment = null;
            DataBank.AreaApartment = null;
            DataBank.RoomsApartment = null;
            DataBank.FloorApartment = null;
            DataBank.SqaureApartment = null;

            DataBank.SNMClients = null;
            DataBank.BudgetClients = null;
            DataBank.AreaClients = null;
            DataBank.RoomsClients = null;
            DataBank.FloorClients = null;
            DataBank.SqaureClients = null;
            DataBank.EmailClients = null;
            Close();
        }

        //открытие формы DataBaseClient
        private void button4_Click(object sender, EventArgs e)
        {
            DataBank.TextValueClients = textBox12.Text;
            DataBank.TextAreaClients = textBox13.Text;
            DataBank.TextRoomsClients = textBox14.Text;
            DataBank.TextFloorClients = textBox15.Text;
            DataBank.TextSqaureClients = textBox16.Text;
            Data_Base.DataBaseClients dataBaseClient = new Data_Base.DataBaseClients(this);
            dataBaseClient.Show();
        }
        
        //открытие формы DataBaseRealtors
        private void button1_Click(object sender, EventArgs e)
        {
            Data_Base.DataBaseRealtors dataBaseRealtors = new Data_Base.DataBaseRealtors(this);
            dataBaseRealtors.Show();
        }

        //открытие формы DataBaseApartments
        private void button3_Click(object sender, EventArgs e)
        {
            DataBank.TextValueApartments = textBox5.Text;
            DataBank.TextAreaApartments = textBox6.Text;
            DataBank.TextRoomsApartments = textBox7.Text;
            DataBank.TextFloorApartments = textBox8.Text;
            DataBank.TextSqaureApartments = textBox9.Text;
            Data_Base.DataBaseApartments dataBaseRealtors = new Data_Base.DataBaseApartments(this);
            dataBaseRealtors.Show();
        }

        //загрузка данных риелторов в форму InsertTransactions
        public async void InsertTransaction_Load_ForRealtors(SqlConnection connection, int id)
        {
            SqlCommand getRealtorsInfoCommand = new SqlCommand("SELECT [Surname], [Name], [MiddleName], [PercentFromDeal], [Email] FROM [Realtors] WHERE [Id]=@Id", sqlConnection);
            getRealtorsInfoCommand.Parameters.AddWithValue("Id", id);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await getRealtorsInfoCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    textBox1.Text = Convert.ToString(sqlReader["Surname"]) + " " + Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["MiddleName"]);
                    textBox2.Text = Convert.ToString(sqlReader["PercentFromDeal"]);
                    textBox3.Text = Convert.ToString(sqlReader["Email"]);
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
        
        //загрузка данных клиентов в форму InsertTransactions
        public async void InsertTransaction_Load_ForClients(SqlConnection connection, int id)
        {
            SqlCommand getClientsInfoCommand = new SqlCommand("SELECT [Surname], [Name], [MiddleName], [Budget], [Area], [Rooms], [Floor], [Sqaure], [Email] FROM [Clients] WHERE [Id]=@Id", sqlConnection);
            getClientsInfoCommand.Parameters.AddWithValue("Id", id);

            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await getClientsInfoCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    textBox4.Text = Convert.ToString(sqlReader["Surname"]) + " " + Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["MiddleName"]);
                    textBox5.Text = Convert.ToString(sqlReader["Budget"]);
                    textBox6.Text = Convert.ToString(sqlReader["Area"]);
                    textBox7.Text = Convert.ToString(sqlReader["Rooms"]);
                    textBox8.Text = Convert.ToString(sqlReader["Floor"]);
                    textBox9.Text = Convert.ToString(sqlReader["Sqaure"]);
                    textBox10.Text = Convert.ToString(sqlReader["Email"]);
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
    
        //загрузка данных квартир в форму InsertTransactions
        public async void InsertTransaction_Load_ForApartments(SqlConnection connection, int id)
        {
            SqlCommand getApartmentsInfoCommand = new SqlCommand("SELECT [Address], [Value], [Area], [Rooms], [Floor], [Sqaure] FROM [Apartments] WHERE [Id]=@Id", sqlConnection);
            getApartmentsInfoCommand.Parameters.AddWithValue("Id", id);
            SqlDataReader sqlReader = null;
            try
            {
                sqlReader = await getApartmentsInfoCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    textBox11.Text = Convert.ToString(sqlReader["Address"]);
                    textBox12.Text = Convert.ToString(sqlReader["Value"]);
                    textBox13.Text = Convert.ToString(sqlReader["Area"]);
                    textBox14.Text = Convert.ToString(sqlReader["Rooms"]);
                    textBox15.Text = Convert.ToString(sqlReader["Floor"]);
                    textBox16.Text = Convert.ToString(sqlReader["Sqaure"]);
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
            DataBank.AddressApartment = null;
            DataBank.ValueApartment = null;
            DataBank.AreaApartment = null;
            DataBank.RoomsApartment = null;
            DataBank.FloorApartment = null;
            DataBank.SqaureApartment = null;

            DataBank.SNMClients = null;
            DataBank.BudgetClients = null;
            DataBank.AreaClients = null;
            DataBank.RoomsClients = null;
            DataBank.FloorClients = null;
            DataBank.SqaureClients = null;
            DataBank.EmailClients = null;
            SqlCommand insertTransactionsCommand = new SqlCommand("INSERT INTO [Transactions] (SNMRealtors, PercentFromDeal, EmailRealtors, SNMClients, BudgetClients, AreaClients, RoomsClients, FloorClients, SqaureClients, EmailClients, AddressApartments, ValueApartments, AreaApartments, RoomsApartments, FloorApartments, SqaureApartments, DateTransactions, ValueTransactions, ValueRealtor, ValueFirm) " +
                "VALUES(@SNMRealtors, @PercentFromDeal, @EmailRealtors, @SNMClients, @BudgetClients, @AreaClients, @RoomsClients, @FloorClients, @SqaureClients, @EmailClients, @AddressApartments, @ValueApartments, @AreaApartments, @RoomsApartments, @FloorApartments, @SqaureApartments, @DateTransactions, @ValueTransactions, @ValueRealtor, @ValueFirm)", sqlConnection);
            try
            {
                insertTransactionsCommand.Parameters.AddWithValue("SNMRealtors", Convert.ToString(textBox1.Text));
                insertTransactionsCommand.Parameters.AddWithValue("PercentFromDeal", Convert.ToInt32(textBox2.Text));
                insertTransactionsCommand.Parameters.AddWithValue("EmailRealtors", Convert.ToString(textBox3.Text));
                insertTransactionsCommand.Parameters.AddWithValue("SNMClients", Convert.ToString(textBox4.Text));
                insertTransactionsCommand.Parameters.AddWithValue("BudgetClients", Convert.ToDecimal(textBox5.Text));
                insertTransactionsCommand.Parameters.AddWithValue("AreaClients", Convert.ToString(textBox6.Text));
                insertTransactionsCommand.Parameters.AddWithValue("RoomsClients", Convert.ToInt32(textBox7.Text));
                insertTransactionsCommand.Parameters.AddWithValue("FloorClients", Convert.ToInt32(textBox8.Text));
                insertTransactionsCommand.Parameters.AddWithValue("SqaureClients", Convert.ToDecimal(textBox9.Text));
                insertTransactionsCommand.Parameters.AddWithValue("EmailClients", Convert.ToString(textBox10.Text));
                insertTransactionsCommand.Parameters.AddWithValue("AddressApartments", Convert.ToString(textBox11.Text));
                insertTransactionsCommand.Parameters.AddWithValue("ValueApartments", Convert.ToDecimal(textBox12.Text));
                insertTransactionsCommand.Parameters.AddWithValue("AreaApartments", Convert.ToString(textBox13.Text));
                insertTransactionsCommand.Parameters.AddWithValue("RoomsApartments", Convert.ToInt32(textBox14.Text));
                insertTransactionsCommand.Parameters.AddWithValue("FloorApartments", Convert.ToInt32(textBox15.Text));
                insertTransactionsCommand.Parameters.AddWithValue("SqaureApartments", Convert.ToDecimal(textBox16.Text));
                insertTransactionsCommand.Parameters.AddWithValue("DateTransactions", Convert.ToString(dateTimePicker1.Text));
            }
            catch
            {
                MessageBox.Show("Вы не выбрали риелтора, клиента или квартиру!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            int percentAgency = 3;
            int percentRealtor = 0;
            decimal valueApartment = 0;
            decimal valueTransaction = 0;
            decimal valueReltor = 0;
            decimal valueFirm = 0;
            try
            {
                percentRealtor = Convert.ToInt32(textBox2.Text);
                valueApartment = Convert.ToDecimal(textBox12.Text);
                valueTransaction = ((valueApartment * (percentRealtor + percentAgency)) / 100) + valueApartment;
                valueReltor = (valueApartment * percentRealtor) / 100;
                valueFirm = (valueApartment * percentAgency) / 100;
                insertTransactionsCommand.Parameters.AddWithValue("ValueTransactions", valueTransaction);
                insertTransactionsCommand.Parameters.AddWithValue("ValueRealtor", valueReltor);
                insertTransactionsCommand.Parameters.AddWithValue("ValueFirm", valueFirm);
                await insertTransactionsCommand.ExecuteNonQueryAsync();//выполняет действие, при этом не возвращая никаких данных
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
                int percentAgency = 3;
                int percentRealtor = 0;
                decimal valueApartment = 0;
                try
                {
                    percentRealtor = Convert.ToInt32(textBox2.Text);
                    valueApartment = Convert.ToDecimal(textBox12.Text);
                }
                catch
                {
                    MessageBox.Show("Пожалуйста, выберите риелтора, клиента и квартиру!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                decimal valueTransaction = 0;
                valueTransaction = ((valueApartment * (percentRealtor + percentAgency)) / 100) + valueApartment;
                label19.Text = Convert.ToString(valueTransaction) + " рублей.";
        }

        private void InsertTransactions_Load(object sender, EventArgs e)
        {
            textBox11.Text = DataBank.AddressApartment;
            textBox12.Text = DataBank.ValueApartment;
            textBox13.Text = DataBank.AreaApartment;
            textBox14.Text = DataBank.RoomsApartment;
            textBox15.Text = DataBank.FloorApartment;
            textBox16.Text = DataBank.SqaureApartment;

            textBox4.Text = DataBank.SNMClients;
            textBox5.Text = DataBank.BudgetClients;
            textBox6.Text = DataBank.AreaClients;
            textBox7.Text = DataBank.RoomsClients;
            textBox8.Text = DataBank.FloorClients;
            textBox9.Text = DataBank.SqaureClients;
            textBox10.Text = DataBank.EmailClients;
        }
    }
}
