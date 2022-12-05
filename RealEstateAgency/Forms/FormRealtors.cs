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
    public partial class FormRealtors : Form
    {
        private ListViewColumnSorter lvwColumnSorter;
        SqlConnection sqlConnection;
        private List<string[]> rows = new List<string[]>();
        private List<string[]> filteretedList = null;
        private List<string[]> filteretedListName = null;
        private List<string[]> filteretedListPercent = null;
        public FormRealtors()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
        }

        private async void FormRealtors_Load(object sender, EventArgs e)
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
                listView1.Columns[listView1.Columns.Count - 1].Width = 200;
                listView1.Columns.Add("Фамилия");
                listView1.Columns[listView1.Columns.Count - 1].Width = 200;
                listView1.Columns.Add("Имя");
                listView1.Columns[listView1.Columns.Count - 1].Width = 200;
                listView1.Columns.Add("Отчество");
                listView1.Columns[listView1.Columns.Count - 1].Width = 200;
                listView1.Columns.Add("Процент от сделки");
                listView1.Columns[listView1.Columns.Count - 1].Width = 200;
                listView1.Columns.Add("Email");
                listView1.Columns[listView1.Columns.Count - 1].Width = 325;
                await LoadRealtorsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormRealtors_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

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
                    //ListViewItem item = new ListViewItem(new string[]
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

        private void listView1_ColumnWidthChanging_1(object sender, ColumnWidthChangingEventArgs e)
        {
            e.Cancel = true;
            e.NewWidth = listView1.Columns[e.ColumnIndex].Width;
        }


        private async void btnUpdate_Click(object sender, EventArgs e)
        {
            rows.Clear();
            await LoadRealtorsAsync();
        }
        //Добавление
        private void btnEdit_Click(object sender, EventArgs e)
        {
            Forms_of_action.InsertRealtors insert = new Forms_of_action.InsertRealtors(sqlConnection);
            insert.Show();
        }
        //Удаление
        private async void btnDelete_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult res = MessageBox.Show("Вы действительно хотите удалить выделенного риелтора?", "Удаление риелтора", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                switch (res)
                {
                    case DialogResult.OK:
                        SqlCommand deleteClientsCommand = new SqlCommand("DELETE FROM [Realtors] WHERE [Id]=@Id", sqlConnection);
                        deleteClientsCommand.Parameters.AddWithValue("Id", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                        try
                        {
                            await deleteClientsCommand.ExecuteNonQueryAsync();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        rows.Clear();
                        await LoadRealtorsAsync();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Редактирование
        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Forms_of_action.UpdateRealtors update = new Forms_of_action.UpdateRealtors(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                update.Show();
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        

        private void btnInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Функции данной вкладки:\n" + "\n" + "Добавить - добавляете нового риелтора в базу данных.\n" + "Удалить - удаление риелтора из базы данных.\n" +
    "Редактировать - изменение данных риелтора.\n" + "Обновить - обновляет базу данных риелтора.\n" + "Также присутствует строка поиска, где можете найти нужного вам риелтора по ФИО.", "О вкладке риелторы:", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            filteretedListName = rows.Where((x) =>
            (x[1].ToLower().Contains(textBox2.Text.ToLower()) || x[2].ToLower().Contains(textBox2.Text.ToLower()) || x[3].ToLower().Contains(textBox2.Text.ToLower())) && x[4].ToLower().Contains(textBox3.Text.ToLower())).ToList();
            RefreshList(filteretedListName);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            filteretedListPercent = rows.Where((x) =>
            (x[1].ToLower().Contains(textBox2.Text.ToLower()) || x[2].ToLower().Contains(textBox2.Text.ToLower()) || x[3].ToLower().Contains(textBox2.Text.ToLower())) && x[4].ToLower().Contains(textBox3.Text.ToLower())).ToList();
            RefreshList(filteretedListPercent);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox3.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Forms_of_action.InsertTransactionsClients insert = new Forms_of_action.InsertTransactionsClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                insert.Show();
                insert.InsertTransaction_Load_ForRealtors(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
            }
            else
            {
                MessageBox.Show("Ни один клиент не выбран!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
