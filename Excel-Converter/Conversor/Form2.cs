using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text;

namespace Conversor
{
    public partial class Form2 : Form
    {
        string NomePadrao = "";
        int NumeroColunas = 0;
        int ColunaAtual = 0;
        bool Botao1Clicado = false;
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(this.textBox1.Text == "")
            {
                Botao1Clicado = false;
            }
            //this.richTextBox1.Clear();
            this.button1.Enabled = true;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(this.textBox1.Text != "")
            {
                Botao1Clicado = true;
            }
            
            
            if(this.textBox1.Text == "" || this.textBox1.Text.Length > 100)
            {
                MessageBox.Show("Nome Inválido!!!");
            }
            else
            {
                NomePadrao = this.textBox1.Text;
            }
            
            
        }

        private void label2_Click(object sender, EventArgs e)
        {
            
            /*if(NumeroColunas > 0 && ColunaAtual <= NumeroColunas)
            {
                ColunaAtual++;
                label2.Text = "Coluna " + ColunaAtual.ToString();
            }
            if(Nome)*/
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.richTextBox1.Clear();
            this.button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if(int.TryParse(this.textBox2.Text, out NumeroColunas) && NumeroColunas > 0)
            {
                this.richTextBox1.AppendText("Total de Colunas: " + NumeroColunas + "\n");
                for (int i = 1; i <= NumeroColunas; i++)
                {
                    this.richTextBox1.AppendText("Coluna " + i + ": " + "\n");
                }
            }
            else
            {
                MessageBox.Show("VALOR INVÁLIDO!!!");
            }
            
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            /*if (!(int.TryParse(this.textBox2.Text, out NumeroColunas) && NumeroColunas > 0 || this.button2.Enabled == true)) 
            {
                richTextBox1.Clear();
            }*/
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!Botao1Clicado)
            {
                MessageBox.Show("Clique no Botão Confirmar para validar o nome do Arquivo Padrão!!");
            }

            if (NomePadrao != "" && (int.TryParse(this.textBox2.Text, out NumeroColunas) && NumeroColunas > 0))
            {
                string initial = System.Windows.Forms.Application.StartupPath;
                
                if (!Directory.Exists(Path.GetFullPath(Path.Combine(initial, "Padroes"))))
                {
                    
                    Directory.CreateDirectory(Path.GetFullPath(Path.Combine(initial, "Padroes")));
                    
                }
                string texto = this.richTextBox1.Text.ToString();
                int contador = 0;
                bool EspaçoSeguinte = false;
                
                foreach (char s in this.richTextBox1.Text)
                {
                    
                    if(EspaçoSeguinte == true && s == ' ')
                    {
                        contador++;
                        EspaçoSeguinte = false;
                        
                    }
                        
                    if (s == ':')
                    {
                        
                        EspaçoSeguinte = true;
                    }
                    
                }

                if(contador != NumeroColunas +1)
                {
                    MessageBox.Show("Preenchimento Inválido");
                    return;
                }
                
                string path = Path.GetFullPath(Path.Combine(initial, "Padroes",NomePadrao + ".txt"));
                using (StreamWriter writer = new StreamWriter(path))
                {
                    writer.WriteLine("");
                    foreach(char s in this.richTextBox1.Text)
                    {
                        writer.Write(s);
                    }
                    writer.Close();
                }
                MessageBox.Show("Padrão criado com sucesso!!");   
            }
            else
            {
                MessageBox.Show("Favor colocar um Nome de arquivo e Número de Colunas Válido!!");
            }
        }    
    }
}
