using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Conversor
{
    public partial class Form4 : Form
    {
        string PastaSaida = "";
        public Form4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(PastaSaida == "")
            {
                MessageBox.Show("Selecione uma pasta de Saída!!!");
            }
            else
            {
                cboSheet f1 = new cboSheet(PastaSaida);
                f1.ShowDialog();
            }
          
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.ShowNewFolderButton = true;
            DialogResult result = folder.ShowDialog();
            if (result == DialogResult.OK)
            {
                PastaSaida = folder.SelectedPath;
                if(PastaSaida == "")
                {
                    MessageBox.Show("Favor Escolher uma pasta de Saída");
                }
                else
                {
                    MessageBox.Show("Pasta de Saída atribuida com Sucesso!");
                }
                Environment.SpecialFolder root = folder.RootFolder;
            }


      
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (PastaSaida == "")
            {
                MessageBox.Show("Selecione uma pasta de Saída!!!");
            }
            else
            {
                ConversaoEmMassa conversao = new ConversaoEmMassa(PastaSaida);
                conversao.ShowDialog();
            }  
        }

      

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Close();
            return;
        }
    }
}
