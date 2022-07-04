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
    public partial class Insert : Form
    {
        private SqlConnection SqlConnection = null;
        public Insert(SqlConnection connection)
        {
            InitializeComponent();
            SqlConnection = connection;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SqlCommand InsertStudentCommand = new SqlCommand("INSERT INTO [Student] (nameStud, course, nameOrg, nameDirPrc, postDir, date, telephone) VALUES(@nameStud, @course, @nameOrg, @nameDirPrc, @postDir, @date, @telephone)", SqlConnection);

            InsertStudentCommand.Parameters.AddWithValue("nameStud", textBox1.Text);
            InsertStudentCommand.Parameters.AddWithValue("course", textBox2.Text);
            InsertStudentCommand.Parameters.AddWithValue("nameOrg", textBox3.Text);
            InsertStudentCommand.Parameters.AddWithValue("nameDirPrc", textBox4.Text);
            InsertStudentCommand.Parameters.AddWithValue("postDir", textBox5.Text);
            InsertStudentCommand.Parameters.AddWithValue("date", textBox6.Text);
            InsertStudentCommand.Parameters.AddWithValue("telephone", textBox7.Text);

            try
            {
                await InsertStudentCommand.ExecuteNonQueryAsync();
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
