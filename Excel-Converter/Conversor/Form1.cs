using System;
using System.IO;
using System.Windows.Forms;
using System.Data;
using ExcelDataReader;
using System.Security.Authentication.ExtendedProtection.Configuration;
using System.Windows.Forms.VisualStyles;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Conversor
{
    
    public partial class cboSheet : Form
    {
        
        string Padrao = "";
        string Arquivo = "";
        string PastaSaida;
        List<int> datas = new List<int>();
        
        
        public cboSheet()
        {
            
            InitializeComponent();
        }

        public cboSheet(string PastaSaida)
        {
            this.PastaSaida = PastaSaida;
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //OpenFile();
           // button1_Click(sender, e);
           // openButton_Click(sender, e);
           // button2_Click(sender, e);
            //button3_Click(sender, e);
            /*Excel excel = new Excel(Arquivo, Padrao, 1);
            bool concluido = excel.PreencherMatriz();
            if (concluido)
            {
                MessageBox.Show("Conversão concluída!!");
            }

            excel.Close();*/


        }

        public void OpenFile()
        {
            //Excel excel = new Excel("C:/Users/Thalles/Desktop/Gravadora/PASTA A - RGE DIGITAL ARTISTICO EMP ERALDO PONTO/RGE PONTO - 1o tri 2013 - 49_49503DPF.XLS", 1);
            //string [,]Matriz = excel.PreencherMatriz();

            
           // MessageBox.Show(excel.ReadCells(0, 0));
            
            
            //MessageBox.Show(excel.PreencherMatriz());
            
            
        }

        DataSet result;
        private void openButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog of = new OpenFileDialog() /*{ Filter = "Excel Workbook|*.xlsx|Excel Work Book 97-2019|*.xls", ValidateNames = true }*/)
            {
                of.Filter = "xls files (*.xls)|*.xls|xlsx (*.xlsx)|*.xlsx";
                if (of.ShowDialog() == DialogResult.OK)
                {
                    //MessageBox.Show(of.FileName);
                    Arquivo = of.FileName;
                    this.textBox1.Text = Arquivo;
                }
                else
                {
                    Arquivo = "";
                }
            }
        }




        

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //dataGridView.DataSource = result.Tables[cbmSheet.SelectedIndex];
            
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog of = new OpenFileDialog())
            {
                string initial = System.Windows.Forms.Application.StartupPath;
                of.InitialDirectory = Path.GetFullPath(Path.Combine(initial, "Padroes"));

                if (!Directory.Exists(of.InitialDirectory))
                {
                    //MessageBox.Show("Nao existe o diretorio inicial");
                    Directory.CreateDirectory(of.InitialDirectory);
                }
                of.InitialDirectory = Path.GetFullPath(of.InitialDirectory);
                of.RestoreDirectory = true;
                of.Filter = "txt files (*.txt)|*.txt";

                //openFileDialog.FilterIndex = 2;

                if (of.ShowDialog() == DialogResult.OK)
                {
                    //MessageBox.Show(of.FileName);
                    Padrao = of.FileName;
                    this.textBox2.Text = this.Padrao;
                }
                else
                {
                    Padrao = "";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool ConversaoEmMassa = false;
            if (Arquivo == "" || Padrao == "")
            {
                MessageBox.Show("Favor Selecionar um Arquivo para ser Convertido e também selecionar um Padrão de Conversao");
            }
            else
            {
                Excel excel = new Excel(Arquivo, Padrao, 1, PastaSaida, ConversaoEmMassa, datas);
                bool concluido = excel.PreencherMatriz();
                if (concluido)
                {
                    MessageBox.Show("Conversão concluída!!");
                    
                }

                excel.Close();
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();

            f2.ShowDialog();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.textBox1.ReadOnly = true;
            //this.button1.Enabled = true;
            // this.button1.Enabled = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.textBox2.ReadOnly = true;
            //this.button2.Enabled = true;
            
        }

        private void cboSheet_Load(object sender, EventArgs e)
        {

        }

    

        private void button3_Click_1(object sender, EventArgs e)
        {   
            bool inteiro = int.TryParse(textBox3.Text.ToString(), out int valor);
            if (inteiro)
            {
                datas.Add(valor);
                /*foreach(int x in datas)
                {
                    MessageBox.Show(x.ToString());
                }*/              
            }
            else
            {
                MessageBox.Show("Valor Inválido!!");
            }
            textBox3.Text = string.Empty;
        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}

