using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Practice
{
    public partial class UPDATE : Form
    {
        private SqlConnection sqlConnection = null;
        private int IdSt;
        public UPDATE(SqlConnection connection, int IdSt)
        {
            InitializeComponent();
            sqlConnection = connection;
            this.IdSt = IdSt;
        }

        private async void UPDATE_Load(object sender, EventArgs e)
        {
            SqlCommand getStudentInfoCommand = new SqlCommand("SELECT [nameStud], [course], [nameOrg], [nameDirPrc], [postDir], [date], [telephone], [address] FROM [Student] WHERE [IdSt]=@IdSt", sqlConnection);
            getStudentInfoCommand.Parameters.AddWithValue("IdSt", IdSt);
            SqlDataReader sqlDataReader = null;

            try
            {
                sqlDataReader = await getStudentInfoCommand.ExecuteReaderAsync();
                while (await sqlDataReader.ReadAsync())
                {
                    textBox1.Text = Convert.ToString(sqlDataReader["nameStud"]);
                    textBox2.Text = Convert.ToString(sqlDataReader["course"]);
                    textBox3.Text = Convert.ToString(sqlDataReader["nameOrg"]);
                    textBox4.Text = Convert.ToString(sqlDataReader["nameDirPrc"]);
                    textBox5.Text = Convert.ToString(sqlDataReader["postDir"]);
                    textBox6.Text = Convert.ToString(sqlDataReader["date"]);
                    textBox7.Text = Convert.ToString(sqlDataReader["telephone"]);
                    textBox9.Text = Convert.ToString(sqlDataReader["address"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            }
            finally
            {
                if (sqlDataReader != null && !sqlDataReader.IsClosed)
                {
                    sqlDataReader.Close();
                }
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SqlCommand update = new SqlCommand("UPDATE [Student] SET [nameStud]=@nameStud,[course]=@course, [nameOrg]=@nameOrg, [nameDirPrc]=@nameDirPrc, [postDir]=@postDir, [date]=@date, [telephone]=@telephone, [address]=@address WHERE [IdSt]=@IdSt", sqlConnection);

            update.Parameters.AddWithValue("IdSt", IdSt);
            update.Parameters.AddWithValue("nameStud", textBox1.Text);
            update.Parameters.AddWithValue("course", textBox2.Text);
            update.Parameters.AddWithValue("nameOrg", textBox3.Text);
            update.Parameters.AddWithValue("nameDirPrc", textBox4.Text);
            update.Parameters.AddWithValue("postDir", textBox5.Text);
            update.Parameters.AddWithValue("date", textBox6.Text);
            update.Parameters.AddWithValue("telephone", textBox7.Text);
            update.Parameters.AddWithValue("address", textBox9.Text);

            try
            {
                await update.ExecuteNonQueryAsync();
                Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }
}
