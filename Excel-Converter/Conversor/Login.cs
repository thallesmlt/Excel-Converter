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
using System.Data;

namespace Conversor
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private bool isValid()
        {
         if(textBox1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Entre com um nome válido!!");
                return false;
            }
            else if(textBox1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Entre com um nome válido!!");
                return false;
            }
           return true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (isValid())
            {
                using (SqlConnection conn = new SqlConnection(@"Data Source=(LocalDb)\MSSQLLocalDB;AttachDBFilename=|DataDirectory|\Database1.mdf;Integrated Security=True"))
                {
                    string query = "SELECT * FROM Login WHERE Username =  '" + textBox1.Text.Trim() + "' AND Password = '" + textBox2.Text.Trim() + "'";
                    SqlDataAdapter sda = new SqlDataAdapter(query, conn);
                    DataTable dta = new DataTable();
                    sda.Fill(dta);
                    if(dta.Rows.Count == 1 || dta.Rows.Count == 2)
                    {
                        //Application.Run(new Form4());
                        Form4 form4 = new Form4();
                        
                        form4.Show();
                        this.Hide();

                    }
                    else
                    {
                        MessageBox.Show("Wrong Username or Password");
                    }
                }
            }
           
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.PasswordChar = '*';
            textBox2.MaxLength = 15;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.MaxLength = 15;
        }
    }
}
