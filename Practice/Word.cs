using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Practice
{
    public partial class Word : Form
    {
        private SqlConnection sqlConnection = null;
        private int IdSt;
        public Word(SqlConnection connection, int IdSt)
        {
            InitializeComponent();
            sqlConnection = connection;
            this.IdSt = IdSt;
        }

        private async void Word_Load(object sender, EventArgs e)
        {
            SqlCommand getStudentInfoCommand = new SqlCommand("SELECT [nameStud], [course], [nameOrg], [nameDirPrc], [postDir], [date], [telephone] FROM [Student] WHERE [IdSt]=@IdSt", sqlConnection);
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

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            SqlCommand update = new SqlCommand("UPDATE [Student] SET [nameStud]=@nameStud,[course]=@course, [nameOrg]=@nameOrg, [nameDirPrc]=@nameDirPrc, [postDir]=@postDir, [date]=@date, [telephone]=@telephone WHERE [IdSt]=@IdSt", sqlConnection);

            string[] tag = { "NAMEORG", "FIORUKINIT", "FIOSTUDINIT", "FIORUK", "ADDRESS", "JOBRUK" };
            string[] zam = { textBox3.Text, textBox4.Text, textBox1.Text, textBox4.Text, "Адрес", textBox5.Text, "INN", "OGRN", "Director" };

            try
            {
                string name = ("NewDoc" + ".docx");
                if (File.Exists(name) == false)
                    File.Copy("obrazec.docx", name);
                var path = Path.GetFullPath(name);
                Console.WriteLine(path);
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                word.Visible = true;
                //Document doc = word.Documents.Open(path);
                word.Documents.Open(path);
                for (int i = 0; i < tag.Length; i++)
                {
                    Find findObject = word.Selection.Find;
                    //findObject.ClearFormatting();
                    findObject.Text = tag[i];
                    //findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = zam[i];

                    Object wrap = WdFindWrap.wdFindContinue;
                    Object replace = WdReplace.wdReplaceAll;
                    findObject.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: Type.Missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: Type.Missing, Replace: replace);
                }
                word.ActiveDocument.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            }
        }
    }
}
