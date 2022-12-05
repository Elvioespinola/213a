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
using SortOrder = System.Windows.Forms.SortOrder;

namespace RealEstateAgency.Forms
{
    public partial class FormClients : Form
    {
        private ListViewColumnSorter lvwColumnSorter;
        SqlConnection sqlConnection;
        private List<string[]> rows = new List<string[]>();
        private List<string[]> filteretedList = null;
        private List<string[]> filteretedListValue = null;
        private List<string[]> filteretedListArea = null;
        private List<string[]> filteretedListRooms = null;
        private List<string[]> filteretedListFloor = null;
        private List<string[]> filteretedListSqaure = null;

        public FormClients()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
        }

        private async void FormClients_Load(object sender, EventArgs e)
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
                listView1.Columns[listView1.Columns.Count - 1].Width = 80;
                listView1.Columns.Add("Фамилия");
                listView1.Columns[listView1.Columns.Count - 1].Width = 125;
                listView1.Columns.Add("Имя");
                listView1.Columns[listView1.Columns.Count - 1].Width = 100;
                listView1.Columns.Add("Отчество");
                listView1.Columns[listView1.Columns.Count - 1].Width = 150;
                listView1.Columns.Add("Бюджет");
                listView1.Columns[listView1.Columns.Count - 1].Width = 125;
                listView1.Columns.Add("Район");
                listView1.Columns[listView1.Columns.Count - 1].Width = 175;
                listView1.Columns.Add("Кол-во комнат");
                listView1.Columns[listView1.Columns.Count - 1].Width = 150;
                listView1.Columns.Add("Этаж");
                listView1.Columns[listView1.Columns.Count - 1].Width = 80;
                listView1.Columns.Add("Площадь (кв/м)");
                listView1.Columns[listView1.Columns.Count - 1].Width = 100;
                listView1.Columns.Add("E-mail");
                listView1.Columns[listView1.Columns.Count - 1].Width = 300;
                await LoadClientsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormClients_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

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
                    //ListViewItem item = new ListViewItem(new string[]
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

        private void listView1_ColumnWidthChanging_1(object sender, ColumnWidthChangingEventArgs e)
        {
            e.Cancel = true;
            e.NewWidth = listView1.Columns[e.ColumnIndex].Width;
        }

        private async void btnUpdate_Click(object sender, EventArgs e)
        {
            rows.Clear();
            await LoadClientsAsync();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            Forms_of_action.InsertClients insert = new Forms_of_action.InsertClients(sqlConnection);
            insert.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Forms_of_action.UpdateClients update = new Forms_of_action.UpdateClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                update.Show();
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnDelete_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult res = MessageBox.Show("Вы действительно хотите удалить выделенного клиента?", "Удаление клиента", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                switch (res)
                {
                    case DialogResult.OK:
                        SqlCommand deleteClientsCommand = new SqlCommand("DELETE FROM [Clients] WHERE [Id]=@Id", sqlConnection);
                        deleteClientsCommand.Parameters.AddWithValue("Id", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                        try
                        {
                            await deleteClientsCommand.ExecuteNonQueryAsync();
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        rows.Clear();
                        await LoadClientsAsync();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (Char.IsDigit(e.KeyChar) == false) return;
        //    if (e.KeyChar == Convert.ToChar(Keys.Back)) return;
        //    e.Handled = true;
        //    textBox1.Clear();
        //    MessageBox.Show("Неправильно введены данные. Пожалуйста, введите ФИО клиента!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //}

        private void btnInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Функции данной вкладки:\n" + "\n" + "Добавить - добавляете нового клиента в базу данных.\n" + "Удалить - удаление клиента из базы данных.\n" +
                "Редактировать - изменение данных клиента.\n"+ "Обновить - обновляет базу данных клиента.\n" + "Также присутствует строка поиска, где можете найти нужного вам клиента по ФИО.", "О вкладке клиенты:", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //сортировка по столбцам
        private void listView1_ColumnClick_1(object sender, ColumnClickEventArgs e)
        {
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filteretedListValue = rows.Where((x) =>
            x[4].ToLower().Contains(textBox2.Text.ToLower()) && x[5].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox4.Text.ToLower()) && x[7].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListValue);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            filteretedListArea = rows.Where((x) =>
            x[4].ToLower().Contains(textBox2.Text.ToLower()) && x[5].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox4.Text.ToLower()) && x[7].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListArea);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            filteretedListRooms = rows.Where((x) =>
            x[4].ToLower().Contains(textBox2.Text.ToLower()) && x[5].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox4.Text.ToLower()) && x[7].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListRooms);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            filteretedListFloor = rows.Where((x) =>
            x[4].ToLower().Contains(textBox2.Text.ToLower()) && x[5].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox4.Text.ToLower()) && x[7].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListFloor);
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            filteretedListSqaure = rows.Where((x) =>
            x[4].ToLower().Contains(textBox2.Text.ToLower()) && x[5].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[6].ToLower().Contains(textBox4.Text.ToLower()) && x[7].ToLower().Contains(textBox5.Text.ToLower()) &&
            x[8].ToLower().Contains(textBox6.Text.ToLower())).ToList();
            RefreshList(filteretedListSqaure);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Forms_of_action.InsertTransactionsClients insert = new Forms_of_action.InsertTransactionsClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                insert.Show();
                insert.InsertTransaction_Load_ForClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
            }
            else
            {
                MessageBox.Show("Ни один клиент не выбран!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
