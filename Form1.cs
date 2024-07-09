using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration; // Добавьте это пространство имен
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.VisualBasic;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private string connectionString;
        private OleDbConnection connection;

        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add("1");
            comboBox1.Items.Add("2");
            comboBox1.Items.Add("3");
            comboBox1.Items.Add("4");
            comboBox1.Items.Add("5");
            comboBox1.Items.Add("6");
            // Считываем строку подключения из настроек
            connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnectionString"].ConnectionString;
            connection = new OleDbConnection(connectionString); // Инициализируем подключение один раз
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
        }
        string a;

        private void LoadData()
        {
            try
            {
                connection.Open(); // Открываем подключение
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter($"SELECT * FROM {a}", connection); // Используем квадратные скобки для имени таблицы
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            // При возникновении ошибок
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
                connection.Close();
            }
            finally
            {
                connection.Close();
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (n == 1)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE FROM [Вид спорта] WHERE [Идентификатор вида спорта] = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
            if (n == 2)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE FROM Команда WHERE [Идентификатор команды] = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
            if (n == 3)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE FROM Результат WHERE [Номер спортсмена] = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
            if (n == 4)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE FROM Соревнование WHERE [Идентификатор соревнования] = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
            if (n == 5)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE Спортсмен FROM [Номер спортсмена] WHERE  = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
            if (n == 6)
            {
                void DeleteRecord(string id)
                {
                    try
                    {
                        connection.Open();

                        string query = $"DELETE FROM Стадион WHERE [Идентификатор стадиона] = @ID";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", id);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record deleted successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                DeleteRecord(textBox1.Text);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (n == 1)
            {
                void UpdateRecord(string value1, string value2)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE [Вид спорта] SET [Название вида спорта] = @Value1 WHERE [Идентификатор вида спорта] = @Value2";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text);
            }
            if (n == 2)
            {
                void UpdateRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE Команда SET [Название команды] = @Value1, Город = @Value2, [Количество игр] = @Value3, Тренер = @Value4 WHERE [Идентификатор команды] = @Value5";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.Parameters.AddWithValue("@Value2", value5);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 3)
            {
                void UpdateRecord(string value1, string value2, string value3, string value4)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE Результат SET [Дата выступления] = @Value1, [Номер попытки] = @Value2, [Результат попытки] = @Value3 WHERE [Номер спортсмена] = @Value4";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
            if (n == 4)
            {
                void UpdateRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE Соревнование SET [Идентфикатор стадиона] = @Value1, [Идентификатор вида спорта] = @Value2, [Дата начала] = @Value3, [Дата окончания] = @Value4 WHERE [Идентификатор соревнования] = @Value5";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 5)
            {
                void UpdateRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE Спортсмен SET Фамилия = @Value1, Имя = @Value2, Отчество =  @Value3, [Идентификатор команды] = @Value4 WHERE [Номер спортсмена] = @Value5";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.Parameters.AddWithValue("@Value2", value5);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 6)
            {
                void UpdateRecord(string value1, string value2, string value3, string value4)
                {
                    try
                    {
                        connection.Open();
                        string query = "UPDATE Стадион SET Название = @Value1, Адрес = @Value2, Вместимость = @Value3 WHERE [Идентификатор стадиона] = @Value4";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record UPDATED successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                UpdateRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (n == 1)
            {
                void AddRecord(string value1, string value2)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO [Вид спорта] ([Идентификатор вида спорта], [Название вида спорта]) VALUES (@Value1, @Value2)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text);
            }
            if (n == 2)
            {
                void AddRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO Команда ([Идентификатор команды], [Название команды], Город, [Количество игр], Тренер) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.Parameters.AddWithValue("@Value2", value5);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 3)
            {
                void AddRecord(string value1, string value2, string value3, string value4)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO Результат ([Номер спортсмена], [Дата выступления], [Номер попытки], [Результат попытки]) VALUES (@Value1, @Value2, @Value3, @Value4)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
            if (n == 4)
            {
                void AddRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO Соревнование ([Идентификатор соревнования], [Идентфикатор стадиона], [Идентификатор вида спорта], [Дата начала], [Дата окончания]) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 5)
            {
                void AddRecord(string value1, string value2, string value3, string value4, string value5)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO Спортсмен ([Номер спортсмена], Фамилия, Имя, Отчество, [Идентификатор команды]) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.Parameters.AddWithValue("@Value2", value5);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text);
            }
            if (n == 6)
            {
                void AddRecord(string value1, string value2, string value3, string value4)
                {
                    try
                    {
                        connection.Open();
                        string query = "INSERT INTO Стадион ([Идентификатор стадиона], Название, Адрес, Вместимость) VALUES (@Value1, @Value2, @Value3, @Value4)";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@Value1", value1);
                        command.Parameters.AddWithValue("@Value2", value2);
                        command.Parameters.AddWithValue("@Value2", value3);
                        command.Parameters.AddWithValue("@Value2", value4);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Record added successfully!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadData();
                    }
                }
                AddRecord(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            }
        }

        int n;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            n = comboBox1.SelectedIndex + 1;

            if (n == 1)
            {
                a = "[Вид спорта]";
                LoadData();
            }
            if (n == 2)
            {
                a = "[Команда]";
                LoadData();
            }
            if (n == 3)
            {
                a = "[Результат]";
                LoadData();
            }
            if (n == 4)
            {
                a = "[Соревнование]";
                LoadData();
            }
            if (n == 5)
            {
                a = "[Спортсмен]";
                LoadData();
            }
            if (n == 6)
            {
                a = "[Стадион]";
                LoadData();
            }
        }



        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
