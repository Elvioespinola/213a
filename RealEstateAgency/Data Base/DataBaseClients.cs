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
    public partial class DataBaseClients : Form 
    {
        //переменные
        private SqlConnection sqlConnection;
        private List<string[]> rows = new List<string[]>();
        private InsertTransactionsClients insertlink;
        private List<string[]> filteretedListValue = null;
        private List<string[]> filteretedListArea = null;
        private List<string[]> filteretedListRooms = null;
        private List<string[]> filteretedListFloor = null;
        private List<string[]> filteretedListSqaure = null;

        public DataBaseClients(InsertTransactionsClients insertlink)
        {
            InitializeComponent();
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
            this.insertlink = insertlink;
        }

        //работа с интерфейсом
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        //загрузка формы
        private async void DataBaseClients_Load(object sender, EventArgs e)
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
                listView1.Columns.Add("Бюджет");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Район");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Кол-во комнат");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Этаж");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Площадь");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("E-mail");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                await LoadClientsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            textBox1.Text = DataBank.TextValueClients;
            textBox2.Text = DataBank.TextAreaClients;
            textBox3.Text = DataBank.TextRoomsClients;
            textBox4.Text = DataBank.TextFloorClients;
            textBox5.Text = DataBank.TextSqaureClients;
        }

        //событие закрытия sqlConnection при закрытии формы
        private void DataBaseClients_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        //заполнение данными listview
        private async Task LoadClientsAsync()
        {
            SqlDataReader sqlReader = null;
            string[] row = null;
            SqlCommand getClientsCommand = new SqlCommand("SELECT * FROM [Clients]", sqlConnection);
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
                        Convert.ToString(sqlReader["Budget"]),
                        Convert.ToString(sqlReader["Area"]),
                        Convert.ToString(sqlReader["Rooms"]),
                        Convert.ToString(sqlReader["Floor"]),
                        Convert.ToString(sqlReader["Sqaure"]),
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

        //захват панели
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        //создает новую форму, но нужно, чтобы данные заносились уже в активной форме
        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                this.insertlink.InsertTransaction_Load_ForClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                Close();
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтр по бюджету
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            filteretedListValue = rows.Where((x) =>
            x[4].ToLower().Contains(textBox1.Text.ToLower()) && x[5].ToLower().Contains(textBox2.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox3.Text.ToLower()) && x[7].ToLower().Contains(textBox4.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox5.Text.ToLower())).ToList();
            RefreshList(filteretedListValue);
        }

        //фильтр по району
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filteretedListArea = rows.Where((x) =>
            x[4].ToLower().Contains(textBox1.Text.ToLower()) && x[5].ToLower().Contains(textBox2.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox3.Text.ToLower()) && x[7].ToLower().Contains(textBox4.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox5.Text.ToLower())).ToList();
            RefreshList(filteretedListArea);
        }

        //фильтр по кол-ву комнат
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            filteretedListRooms = rows.Where((x) =>
            x[4].ToLower().Contains(textBox1.Text.ToLower()) && x[5].ToLower().Contains(textBox2.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox3.Text.ToLower()) && x[7].ToLower().Contains(textBox4.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox5.Text.ToLower())).ToList();
            RefreshList(filteretedListRooms);
        }

        //фильтр по этажу
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            filteretedListFloor = rows.Where((x) =>
            x[4].ToLower().Contains(textBox1.Text.ToLower()) && x[5].ToLower().Contains(textBox2.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox3.Text.ToLower()) && x[7].ToLower().Contains(textBox4.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox5.Text.ToLower())).ToList();
            RefreshList(filteretedListFloor);
        }

        //фильтр по площади
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            filteretedListSqaure = rows.Where((x) =>
            x[4].ToLower().Contains(textBox1.Text.ToLower()) && x[5].ToLower().Contains(textBox2.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox3.Text.ToLower()) && x[7].ToLower().Contains(textBox4.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox5.Text.ToLower())).ToList();
            RefreshList(filteretedListSqaure);
        }

        //очистка фильтора
        private void button1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
        }
    }
}
