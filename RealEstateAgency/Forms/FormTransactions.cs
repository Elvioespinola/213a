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
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace RealEstateAgency.Forms
{
    public partial class FormTransactions : Form
    {
        SqlConnection sqlConnection;
        private int id;
        private List<string[]> rows = new List<string[]>();
        private List<string[]> filteretedList = null;
        private List<string[]> filteretedListDate = null;
        private List<string[]> filteretedListSNMClient = null;
        private List<string[]> filteretedListSNMRealtor = null;
        private ListViewColumnSorter lvwColumnSorter;        
        public FormTransactions()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
        }

        private async void FormTransactions_Load(object sender, EventArgs e)
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
                listView1.Columns[listView1.Columns.Count - 1].Width = 125;
                listView1.Columns.Add("Дата создания");
                listView1.Columns[listView1.Columns.Count - 1].Width = 210;
                listView1.Columns.Add("Общая сумма");
                listView1.Columns[listView1.Columns.Count - 1].Width = 180;
                listView1.Columns.Add("Покупатель");
                listView1.Columns[listView1.Columns.Count - 1].Width = 500;
                listView1.Columns.Add("Риелтор");
                listView1.Columns[listView1.Columns.Count - 1].Width = 500;
                await LoadTransactionsAsync();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с базой данный!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormTransactions_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private async Task LoadTransactionsAsync()
        {
            SqlDataReader sqlReader = null;
            string[] row = null;
            SqlCommand getClientsCommand = new SqlCommand("SELECT * FROM [Transactions]", sqlConnection);
            try
            {
                sqlReader = await getClientsCommand.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    //ListViewItem item = new ListViewItem(new string[]
                    row = new string[]
                    {
                        Convert.ToString(sqlReader["Id"]),
                        Convert.ToString(sqlReader["DateTransactions"]),
                        Convert.ToString(sqlReader["ValueTransactions"]),
                        Convert.ToString(sqlReader["SNMClients"]),
                        Convert.ToString(sqlReader["SNMRealtors"])
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

        //форма обновления транзакции
        private async void btnUpdate_Click(object sender, EventArgs e)
        {
            rows.Clear();
            await LoadTransactionsAsync();
        }

        //форма добавления заказа
        private void btnAdd_Click(object sender, EventArgs e)
        {
            Forms_of_action.InsertTransactionsClients insert = new Forms_of_action.InsertTransactionsClients(sqlConnection, Convert.ToInt32(listView1.SelectedItems.Count));
            insert.Show();
        }

        //форма удаления заказа
        private async void btnDelete_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult res = MessageBox.Show("Вы действительно хотите удалить выделенного клиента?", "Удаление клиента", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                switch (res)
                {
                    case DialogResult.OK:
                        SqlCommand deleteTransactionsCommand = new SqlCommand("DELETE FROM [Transactions] WHERE [Id]=@Id", sqlConnection);
                        deleteTransactionsCommand.Parameters.AddWithValue("Id", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                        try
                        {
                            await deleteTransactionsCommand.ExecuteNonQueryAsync();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        rows.Clear();
                        await LoadTransactionsAsync();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
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

        //фильтрация
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filteretedListDate = rows.Where((x) =>
            x[1].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower())).ToList();
            RefreshList(filteretedListDate);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            filteretedListSNMClient = rows.Where((x) =>
            x[1].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower())).ToList();
            RefreshList(filteretedListSNMClient);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            filteretedListSNMRealtor = rows.Where((x) =>
            x[1].ToLower().Contains(textBox2.Text.ToLower()) && x[3].ToLower().Contains(textBox3.Text.ToLower()) &&
            x[4].ToLower().Contains(textBox4.Text.ToLower())).ToList();
            RefreshList(filteretedListSNMRealtor);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
        }


        //кнопка экспорта
        private async void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DialogResult res = MessageBox.Show("Вы действительно хотите экспортировать выбранную сделку в MS Word?", "Экспорт сделки", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                switch (res)
                {
                    case DialogResult.OK:
                        SqlCommand getApartmentsInfoCommand = new SqlCommand("SELECT [ValueTransactions], [ValueFirm], [ValueRealtor], [SqaureApartments], [FloorApartments], [ValueApartments], [AreaApartments], [RoomsApartments], [AddressApartments], [EmailClients], [SNMRealtors], [SNMClients], [DateTransactions], [PercentFromDeal], [EmailRealtors], [BudgetClients] FROM [Transactions] WHERE [Id]=@Id", sqlConnection);
                        getApartmentsInfoCommand.Parameters.AddWithValue("Id", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                        SqlDataReader sqlReader = null;
                        try
                        {
                            sqlReader = await getApartmentsInfoCommand.ExecuteReaderAsync();
                            while (await sqlReader.ReadAsync())
                            {
                                try
                                {
                                    var helper = new WordHelper(@"C:\Users\User\source\repos\RealEstateAgency\RealEstateAgency\bin\Debug\transactions\transaction.docx");
                                    var realtor = Convert.ToString(sqlReader["SNMRealtors"]);
                                    var client = Convert.ToString(sqlReader["SNMClients"]);
                                    var idTransaction = Convert.ToString(listView1.SelectedItems[0].SubItems[0].Text);
                                    var date = Convert.ToString(sqlReader["DateTransactions"]);
                                    var percentRealtor = Convert.ToString(sqlReader["PercentFromDeal"]);
                                    var emailRealtor = Convert.ToString(sqlReader["EmailRealtors"]);
                                    var budget = Convert.ToString(sqlReader["BudgetClients"]);
                                    var emailClient = Convert.ToString(sqlReader["EmailClients"]);
                                    var address = Convert.ToString(sqlReader["AddressApartments"]);
                                    var area = Convert.ToString(sqlReader["AreaApartments"]);
                                    var value = Convert.ToString(sqlReader["ValueApartments"]);
                                    var rooms = Convert.ToString(sqlReader["RoomsApartments"]);
                                    var floor = Convert.ToString(sqlReader["FloorApartments"]);
                                    var sqaure = Convert.ToString(sqlReader["SqaureApartments"]);
                                    var valueRealtor = Convert.ToString(sqlReader["ValueRealtor"]);
                                    var valueFirm = Convert.ToString(sqlReader["ValueFirm"]);
                                    var valueTransactions = Convert.ToString(sqlReader["ValueTransactions"]);
                                    var items = new Dictionary<string, string>
                                    {
                                        {"{Realtor}", realtor },
                                        {"{Client}", client },
                                        {"{Id}", idTransaction },
                                        {"{Date}", date },
                                        {"{PercentRealtor}", percentRealtor },
                                        {"{EmailRealtor}", emailRealtor },
                                        {"{Budget}", budget },
                                        {"{EmailClient}", emailClient },
                                        {"{Address}", address },
                                        {"{Area}", area },
                                        {"{Value}", value },
                                        {"{Rooms}", rooms },
                                        {"{Floor}", floor },
                                        {"{Sqaure}", sqaure },
                                        {"{ValueRealtor}", valueRealtor },
                                        {"{ValueFirm}", valueFirm },
                                        {"{ValueTransaction}", valueTransactions },
                                    };
                                    helper.Process(items);
                                    MessageBox.Show("Вы успешно экспортировали сделку в MS Word! Пожалуйста, зайдите в директорию, где у вас хранятся сделки.", "Экспорт сделки", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
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
                        break;
                }
            }
            else
            {
                MessageBox.Show("Ни одна строка не была выделена!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\User\source\repos\RealEstateAgency\RealEstateAgency\bin\Debug\transactions\transaction.docx";
            Cmd(path);
        }

        void Cmd(string line)
        {
            Process.Start(new ProcessStartInfo { FileName = "explorer", Arguments = $"/n, /select, {line }" });
        }
    }
}
