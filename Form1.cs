using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration; // �������� ��� ������������ ����
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
            // ��������� ������ ����������� �� ��������
            connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnectionString"].ConnectionString;
            connection = new OleDbConnection(connectionString); // �������������� ����������� ���� ���
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
                connection.Open(); // ��������� �����������
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter($"SELECT * FROM {a}", connection); // ���������� ���������� ������ ��� ����� �������
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            // ��� ������������� ������
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

                        string query = $"DELETE FROM [��� ������] WHERE [������������� ���� ������] = @ID";
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

                        string query = $"DELETE FROM ������� WHERE [������������� �������] = @ID";
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

                        string query = $"DELETE FROM ��������� WHERE [����� ����������] = @ID";
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

                        string query = $"DELETE FROM ������������ WHERE [������������� ������������] = @ID";
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

                        string query = $"DELETE ��������� FROM [����� ����������] WHERE  = @ID";
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

                        string query = $"DELETE FROM ������� WHERE [������������� ��������] = @ID";
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
                        string query = "UPDATE [��� ������] SET [�������� ���� ������] = @Value1 WHERE [������������� ���� ������] = @Value2";
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
                        string query = "UPDATE ������� SET [�������� �������] = @Value1, ����� = @Value2, [���������� ���] = @Value3, ������ = @Value4 WHERE [������������� �������] = @Value5";
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
                        string query = "UPDATE ��������� SET [���� �����������] = @Value1, [����� �������] = @Value2, [��������� �������] = @Value3 WHERE [����� ����������] = @Value4";
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
                        string query = "UPDATE ������������ SET [������������ ��������] = @Value1, [������������� ���� ������] = @Value2, [���� ������] = @Value3, [���� ���������] = @Value4 WHERE [������������� ������������] = @Value5";
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
                        string query = "UPDATE ��������� SET ������� = @Value1, ��� = @Value2, �������� =  @Value3, [������������� �������] = @Value4 WHERE [����� ����������] = @Value5";
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
                        string query = "UPDATE ������� SET �������� = @Value1, ����� = @Value2, ����������� = @Value3 WHERE [������������� ��������] = @Value4";
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
                        string query = "INSERT INTO [��� ������] ([������������� ���� ������], [�������� ���� ������]) VALUES (@Value1, @Value2)";
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
                        string query = "INSERT INTO ������� ([������������� �������], [�������� �������], �����, [���������� ���], ������) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
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
                        string query = "INSERT INTO ��������� ([����� ����������], [���� �����������], [����� �������], [��������� �������]) VALUES (@Value1, @Value2, @Value3, @Value4)";
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
                        string query = "INSERT INTO ������������ ([������������� ������������], [������������ ��������], [������������� ���� ������], [���� ������], [���� ���������]) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
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
                        string query = "INSERT INTO ��������� ([����� ����������], �������, ���, ��������, [������������� �������]) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5)";
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
                        string query = "INSERT INTO ������� ([������������� ��������], ��������, �����, �����������) VALUES (@Value1, @Value2, @Value3, @Value4)";
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
                a = "[��� ������]";
                LoadData();
            }
            if (n == 2)
            {
                a = "[�������]";
                LoadData();
            }
            if (n == 3)
            {
                a = "[���������]";
                LoadData();
            }
            if (n == 4)
            {
                a = "[������������]";
                LoadData();
            }
            if (n == 5)
            {
                a = "[���������]";
                LoadData();
            }
            if (n == 6)
            {
                a = "[�������]";
                LoadData();
            }
        }



        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
