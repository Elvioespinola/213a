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
using RealEstateAgency.Forms_of_action;

namespace RealEstateAgency.Data_Base
{
    public partial class DataBaseRealtors : Form
    {
        private SqlConnection sqlConnection;
        private List<string[]> rows = new List<string[]>();
        private InsertTransactionsClients insertRealtors;
        private List<string[]> filteretedListSNM = null;
        private List<string[]> filteretedListProcent = null;

        public DataBaseRealtors(InsertTransactionsClients insertlink)
        {
            InitializeComponent();
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
            this.insertRealtors = insertlink;
        }

        //интерфейс
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        //событие прогрузки окна формы DataBaseRealtors
        private async void DataBaseRealtors_Load(object sender, EventArgs e)
        {
            try
            {
                string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\User\source\repos\RealEstateAgency\RealEstateAgency\Database.mdf;Integrated Security=True";
                sqlConnection = new SqlConnection(connectionString);
                await sqlConnection.OpenAsync();
                listView1.GridLines = true;
                listView1.FullRowSelect = true;
                listView1.View = View.Details;
                listView1.Columns.Add("Номер");
                listView1.Columns[listView1.Columns.Count - 1].Width = 100;
                listView1.Columns.Add("Фамилия");
                listView1.Columns[listView1.Columns.Count - 1].Width = 170;
                listView1.Columns.Add("Имя");
                listView1.Columns[listView1.Columns.Count - 1].Width = 125;
                listView1.Columns.Add("Отчество");
                listView1.Columns[listView1.Columns.Count - 1].Width = 150;
                listView1.Columns.Add("Процент от сделки");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("E-mail");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                await LoadRealtorsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        //закрытие окна
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        //захват панели
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        //событие закрытия sqlConnection при закрытии формы
        private void DataBaseRealtors_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        //заполнение листа
        private async Task LoadRealtorsAsync()
        {
            SqlDataReader sqlReader = null;
            string[] row = null;
            SqlCommand getClientsCommand = new SqlCommand("SELECT * FROM [Realtors]", sqlConnection);
            try
            {
                sqlReader = await getClientsCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    row = new string[]
                    {
                        Convert.ToString(sqlReader["Id"]),
                        Convert.ToString(sqlReader["Surname"]),
                        Convert.ToString(sqlReader["Name"]),
                        Convert.ToString(sqlReader["MiddleName"]),
                        Convert.ToString(sqlReader["PercentFromDeal"]),
                        Convert.ToString(sqlReader["Email"])
                    };
                    rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null && !sqlReader.IsClosed)
                {
                    sqlReader.Close();
                }
            }
            RefreshList(rows);
        }

        private void RefreshList(List<string[]> list)
        {
            listView1.Items.Clear();
            foreach (string[] s in list)
            {

                listView1.Items.Add(new ListViewItem(s));
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                this.insertRealtors.InsertTransaction_Load_ForRealtors(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                Close();
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            filteretedListSNM = rows.Where((x) =>
            (x[1].ToLower().Contains(textBox1.Text.ToLower()) || x[2].ToLower().Contains(textBox1.Text.ToLower()) || x[3].ToLower().Contains(textBox1.Text.ToLower())) && x[4].ToLower().Contains(textBox2.Text.ToLower())).ToList();
            RefreshList(filteretedListSNM);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filteretedListProcent = rows.Where((x) =>
            (x[1].ToLower().Contains(textBox1.Text.ToLower()) || x[2].ToLower().Contains(textBox1.Text.ToLower()) || x[3].ToLower().Contains(textBox1.Text.ToLower())) && x[4].ToLower().Contains(textBox2.Text.ToLower())).ToList();
            RefreshList(filteretedListProcent);
        }

        //очистка фильтора
        private void button1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            textBox1.Clear();
            textBox2.Clear();
        }
    }
}
