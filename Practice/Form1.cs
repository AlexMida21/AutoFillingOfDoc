using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Configuration;

namespace Practice
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        //public string link = "https://www.list-org.com/company/4808097";
        
        public string[] spec;
        public string[] UndGrade_IB = { "26.30.16 ", "26.30.19 ", "26.30.3 ", "71.20.9 ", "72.19.2 ", "74.90.32 ", "95.12 ", "62.01 ", "62.02.1 ", "62.02.2 ", "62.09 ", "62.02.3 ", "62.02.4 ", "62.02.9 ", "74.90 ", "80.20 " };
        public string[] UndGrade_PRI = { "62.0 ", "62.01 ", "62.09 ", "63.11 " };
        public string[] UndGrade_BIN = { "62.0 ", "73.20.1 ", "70.22 ", "63.1 ", "63.11.1 ", "62.09 " };
        public string[] MagGrade_IB = { "26.30.16 ", "26.30.19 ", "26.30.3 ", "71.20.9 ", "72.19.2 ", "74.90.32 ", "95.12 ", "62.01 ", "62.02.1 ", "62.02.2 ", "62.02.4 ", "62.02.9 ", "62.03.1 ", "62.09 ", "63.11.1 ", "62.02.3 ", "74.90 ", "80.20 " };
        public string[] MagGrade_PRI = { "62.0 ", "62.09 ", "63.11.1 ", "72.19 ", "62.01 ", "63.11 " };
        public Form1()
        {
            InitializeComponent();
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
            this.studentTableAdapter.Fill(this.websiteDataSet.Student);
            string connectionString = ConfigurationManager.ConnectionStrings["Practice.Properties.Settings.websiteConnectionString"].ConnectionString;
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();

            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.View = View.Details;
            listView1.Columns.Add("ID");
            listView1.Columns.Add("ФИО студента");
            listView1.Columns.Add("Курс");
            listView1.Columns.Add("Организация");
            listView1.Columns.Add("Дата проведения");
            listView1.Columns.Add("Телефон");
            listView1.Columns.Add("Адрес");

            await LoadData();
        }

        private async Task LoadData()
        {
            SqlDataReader dataReader = null;
            SqlCommand getStud = new SqlCommand("SELECT * FROM [Student]", sqlConnection);
            try 
            { 
                dataReader = await getStud.ExecuteReaderAsync();

                while (await dataReader.ReadAsync())
                {
                    ListViewItem item = new ListViewItem(new String[]
                    {
                        Convert.ToString(dataReader["IdSt"]),
                        Convert.ToString(dataReader["nameStud"]),
                        Convert.ToString(dataReader["course"]),
                        Convert.ToString(dataReader["nameOrg"]),
                        Convert.ToString(dataReader["date"]),
                        Convert.ToString(dataReader["telephone"]),
                        Convert.ToString(dataReader["address"])
                    });
                    listView1.Items.Add(item);
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
            finally 
            { 
                if(dataReader != null && !dataReader.IsClosed)
                    dataReader.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {            
            studentTableAdapter.Update(this.websiteDataSet.Student);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        public async void UpdateList_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            await LoadData();
        }

        private void InsertButton_Click(object sender, EventArgs e)
        {
            Insert insert = new Insert(sqlConnection);
            insert.Show();
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                UPDATE update = new UPDATE(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                update.Show();
            }
            else
            {
                MessageBox.Show("Выберите студента из списка", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private async void DeleteButton_Click(object sender, EventArgs e)
        {
            if(listView1.SelectedItems.Count > 0)
            { 
                DialogResult res = MessageBox.Show("Вы уверены что хотите удалить выделенную строку?", "Удаление", MessageBoxButtons.OKCancel);
                if (res == DialogResult.OK)
                {
                    SqlCommand deleteStudent = new SqlCommand("DELETE FROM [Student] WHERE [IdSt]=@IdSt", sqlConnection);
                    deleteStudent.Parameters.AddWithValue("IdSt",Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));

                    try
                    {
                        await deleteStudent.ExecuteNonQueryAsync();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
                    }
                    listView1.Items.Clear();
                    await LoadData();
                }
            }
            else
            {
                MessageBox.Show("Выберите студента из списка", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                Word word = new Word(sqlConnection, Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                word.Show();
            }
            else
            {
                MessageBox.Show("Выберите студента из списка", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private async void toolStripButton1_Click(object sender, EventArgs e)
        {
            string link = "";
            string okved = "";
            if (listView1.SelectedItems.Count > 0)
            {
                if(CourseSwitch.Text!="")
                { 
                     try
                     {
                         SqlCommand getStudentInfoCommand = new SqlCommand("SELECT [link] FROM [Student] WHERE [IdSt]=@IdSt", sqlConnection);
                         getStudentInfoCommand.Parameters.AddWithValue("IdSt", Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text));
                         SqlDataReader sqlDataReader = null;
                         try
                         {
                             sqlDataReader = await getStudentInfoCommand.ExecuteReaderAsync();
                             while (await sqlDataReader.ReadAsync())
                             {
                                 link = Convert.ToString(sqlDataReader["link"]);
                             }
                         }
                         catch (Exception ex)
                         {
                             MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
                         }
                         File.WriteAllText("link.txt", link); //передаём ссылку в Python
                         Process.Start(new ProcessStartInfo("parse.exe") //запускаем код 
                         { CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }); //скрываем консоль
                         okved = File.ReadAllText("okveds.txt");
                         okved = Regex.Replace(okved, "(\\d{2}.\\d{2})", "\n$1");
                         okved = Regex.Replace(okved, "(по коду ОКВЭД ред.2)", "по коду оквэд");
                         okved = Regex.Replace(okved, "([А-Я])", " — $1");
                         string okvmatch = ("");

                         for (int i = 0; i < spec.Length; i++)
                         {
                             if (okved.Contains(spec[i]))
                             {
                                 okvmatch += spec[i] + " ";
                             }
                         }
                         if (okvmatch == (""))
                         {
                            MessageBox.Show("Подходящие ОКВЭД не найдены", "Ошибка", MessageBoxButtons.OK);
                         }


                        DialogResult res = MessageBox.Show("Подходящие ОКВЭД:\n" + okvmatch + "Показать все ОКВЭД этой компании?", "ОКВЭД", MessageBoxButtons.YesNo);
                         if (res == DialogResult.Yes)
                         {
                             try
                             {
                                 MessageBox.Show(okved, "Все ОКВЭД", MessageBoxButtons.OK);
                             }
                             catch (Exception ex)
                             {
                                 MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
                     }
                }
                else
                {
                    MessageBox.Show("Выберите направление подготовки", "Ошибка", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Выберите студента из списка", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private void CourseSwitch_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CourseSwitch.Text)
            {
                case "Бакалавриат ИБ":
                    {
                        spec = UndGrade_IB;
                        break;
                    }
                case "Бакалавриат ПРИ":
                    {
                        spec = UndGrade_PRI;
                        break;
                    }
                case "Бакалавриат БИН":
                    {
                        spec = UndGrade_BIN;
                        break;
                    }
                case "Магистратура ИБ":
                    {
                        spec = MagGrade_IB;
                        break;
                    }
                case "Магистратура ПРИ":
                    {
                        spec = MagGrade_PRI;
                        break;
                    }
            }
        }
    }
}
