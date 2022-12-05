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
    public partial class DataBaseApartments : Form
    {
        //переменные
        private SqlConnection sqlConnection;
        private List<string[]> rows = new List<string[]>();
        private List<string[]> filteretedListForAddress = null;
        private List<string[]> filteretedListValue = null;
        private List<string[]> filteretedListArea = null;
        private List<string[]> filteretedListRooms = null;
        private List<string[]> filteretedListFloor = null;
        private List<string[]> filteretedListSqaure = null;
        private InsertTransactionsClients insertApartments;

        public DataBaseApartments(InsertTransactionsClients insertlink)
        {
            InitializeComponent();
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
            this.insertApartments = insertlink;
        }


        //интерфейс
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        //загрузка формы DataBaseApartments
        private async void DataBaseApartments_Load(object sender, EventArgs e)
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
                listView1.Columns.Add("Квартира");
                listView1.Columns[listView1.Columns.Count - 1].Width = 170;
                listView1.Columns.Add("Цена");
                listView1.Columns[listView1.Columns.Count - 1].Width = 125;
                listView1.Columns.Add("Район");
                listView1.Columns[listView1.Columns.Count - 1].Width = 150;
                listView1.Columns.Add("Кол-во комнат");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Этаж");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                listView1.Columns.Add("Площадь");
                listView1.Columns[listView1.Columns.Count - 1].Width = 250;
                await LoadApartmentsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            textBox2.Text = DataBank.TextValueApartments;
            textBox3.Text = DataBank.TextAreaApartments;
            textBox4.Text = DataBank.TextRoomsApartments;
            textBox5.Text = DataBank.TextFloorApartments;
            textBox6.Text = DataBank.TextSqaureApartments;
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
        private void DataBaseApartments_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        //заполнение листа
        private async Task LoadApartmentsAsync()
        {
            SqlDataReader sqlReader = null;
            string[] row = null;
            SqlCommand getApartmentsCommand = new SqlCommand("SELECT * FROM [Apartments]", sqlConnection);
            try
            {
                sqlReader = await getApartmentsCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    row = new string[]
                    {
                        Convert.ToString(sqlReader["Id"]),
                        Convert.ToString(sqlReader["Address"]),
                        Convert.ToString(sqlReader["Value"]),
                        Convert.ToString(sqlReader["Area"]),
                        Convert.ToString(sqlReader["Rooms"]),
                        Convert.ToString(sqlReader["Floor"]),
                        Convert.ToString(sqlReader["Sqaure"])
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

        //событие кнопки "выбрать из списка"
        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                this.insertApartments.InsertTransaction_Load_ForApartments(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                Close();
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //фильтр по стоимости
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filteretedListValue = rows.Where((x) =>
            x[2].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower()) && x[5].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListValue);
        }

        //фильтр по району
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            filteretedListArea = rows.Where((x) =>
            x[2].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower()) && x[5].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListArea);
        }

        //фильтр по кол-ву комнат
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            filteretedListRooms = rows.Where((x) =>
            x[2].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower()) && x[5].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListRooms);
        }

        //фильтр по этажу
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            filteretedListFloor = rows.Where((x) =>
            x[2].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower()) && x[5].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListFloor);
        }

        //фильтр по площади
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            filteretedListSqaure = rows.Where((x) =>
            x[2].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower()) && x[5].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListSqaure);
        }

        //очистка фильтора
        private void button1_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
        }
    }
}
