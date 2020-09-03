using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.IO;

namespace Conversor
{
    public partial class ConversaoEmMassa : Form
    {
        string PastaEntrada = "";
        string PastaSaida = "";
        string Padrao = "";
        List<int> datas = new List<int>();
        public ConversaoEmMassa()
        {
            InitializeComponent();
        }

        public ConversaoEmMassa(string PastaSaida)
        {
            this.PastaSaida = PastaSaida;
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            bool ConversaoEmMassa = true;
            bool Concluido = true;
            if (PastaEntrada == "" || Padrao == "")
            {
                MessageBox.Show("Favor Selecionar um Arquivo para ser Convertido e também selecionar um Padrão de Conversão");
                return;
            }
            else
            {
                ProgressBarcs progress = new ProgressBarcs(PastaEntrada, PastaSaida, Padrao, ConversaoEmMassa, datas);
                progress.ShowDialog();
                
                
                /*string[] filePaths = Directory.GetFiles(PastaEntrada, "*.xls");
             
                for(int i = 0; i < filePaths.Length -1; i++)
                {
                    //MessageBox.Show(filePaths[i]);
                    Excel excel = new Excel(filePaths[i], Padrao, 1, PastaSaida,ConversaoEmMassa);
                    bool concluido = excel.PreencherMatriz();
                    if (!concluido)
                    {
                        MessageBox.Show("Erro durante a Conversão!! \n  Arquivo: " + filePaths[i]);
                        //MessageBox.Show("Conversão concluída!!");
                    }
                    else
                    {
                        concluido = false;
                    }

                    excel.Close();



                }*/
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.ShowNewFolderButton = true;
            DialogResult result = folder.ShowDialog();
            if (result == DialogResult.OK)
            {
                PastaEntrada = folder.SelectedPath;
                if (PastaEntrada != "")
                {
                    this.textBox1.Text = PastaEntrada;
                    Environment.SpecialFolder root = folder.RootFolder;
                }   
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.textBox1.ReadOnly = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.textBox2.ReadOnly = true;
        }

        private void button2_Click(object sender, EventArgs e)
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

        private void button4_Click(object sender, EventArgs e)
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
    }
}















