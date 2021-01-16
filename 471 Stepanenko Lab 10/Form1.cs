using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _471_Stepanenko_Lab_10
{
    public partial class Form1 : Form
    {
        public static string connectString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", @"C:\Users\USER\source\repos\471 Stepanenko Lab 10\471 Stepanenko Lab 10\database.mdb");
        // вариант 2
        //public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=database.mdb;";

        // поле - ссылка на экземпляр класса OleDbConnection для соединения с БД
        OleDbConnection myConnection;
        public Form1()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetData();
        }

        private void deleteAll_Click(object sender, EventArgs e)
        {
            DeleteAll();
            GetData();
        }

        private void deleteOne_Click(object sender, EventArgs e)
        {
            DeleteOne();
            GetData();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            PushData();
            GetData();
        }

        private void GetData()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string query = "SELECT * from Subjects";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    OleDbDataReader reader = command.ExecuteReader();

                    //label1.Text = command.ExecuteScalar().ToString();
                    listBox1.Items.Clear();
                    while (reader.Read())
                    {
                        listBox1.Items.Add(String.Format("{0} {1}", reader[0].ToString(), reader[1].ToString()));
                    }
                    reader.Close();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void PushData()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string value = textBox1.Text;
                    if (value == "")
                    {
                        MessageBox.Show("Ви не ввели значення запису", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        string query = String.Format("insert into Subjects (subject_name) values (\"{0}\")", value);
                        OleDbCommand command = new OleDbCommand(query, myConnection);
                        command.ExecuteNonQuery();
                    }
                    
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void DeleteOne()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string[] index = listBox1.SelectedItem.ToString().Split(new char[] { ' ', '\n' });
                    string query = String.Format("DELETE FROM Subjects WHERE Код = {0}", index[0]);

                    // создаем объект OleDbCommand для выполнения запроса к БД MS Access
                    OleDbCommand command = new OleDbCommand(query, myConnection);

                    // выполняем запрос к MS Access
                    command.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void DeleteAll(){
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string query = "DELETE FROM Subjects";

                    // создаем объект OleDbCommand для выполнения запроса к БД MS Access
                    OleDbCommand command = new OleDbCommand(query, myConnection);

                    // выполняем запрос к MS Access
                    command.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //лише технічні предмети
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string query = "SELECT * from Subjects where type = \"Технічна\"";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    OleDbDataReader reader = command.ExecuteReader();

                    //label1.Text = command.ExecuteScalar().ToString();
                    listBox1.Items.Clear();
                    while (reader.Read())
                    {
                        listBox1.Items.Add(String.Format("{0} {1}", reader[0].ToString(), reader[1].ToString()));
                    }
                    reader.Close();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
