using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Globalization;



namespace Conversor
{
    class Excel
    {
        public int Colunas;
        public int Linhas;
        public string[,] matriz;
        public string Padrao;
        public string path;
        public List<int> Datas;
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
       
        public string[,] MatrizResultado;
        public int[] Atribuicoes;
        public string PastaSaida;
        public bool ConversaoEmMassa;
        bool VarrerColuna = false;
        char[] charsToTrim = { 'E', '-' };




        public Excel(string path, string Padrao, int Sheet,string PastaSaida,bool ConversaoEmMassa, List<int> Datas)
        {
            //this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            this.path = path;
            _Excel.Range range = ws.UsedRange;
            Linhas = range.Rows.Count;
            Colunas = range.Columns.Count;
            matriz = new string[Linhas + 1,Colunas + 1];
            this.Padrao = Padrao;
            this.PastaSaida = PastaSaida;
            this.ConversaoEmMassa = ConversaoEmMassa;
            this.Datas = Datas;
        }

        public string ReadCells(int i, int j)
        {
            
           i++;
            j++;
            if(ws.Cells[i,j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;  
            }
            else
            {
                return "";
            }
            
        }

        /*public string[, ] */ 
        public bool PreencherMatriz()
        {
            //string outubro = "10 / 51";
            //string outubro2 = "10 / 90";
            if (path.Trim(' ').Length == 0 || Padrao.Trim(' ').Length == 0)
            {
                return false;
            }

          
           // MessageBox.Show(Linhas.ToString());
           // MessageBox.Show(Colunas.ToString());
            for (int i = 1; i < Linhas + 1; i++)
            {
                for(int j = 1; j < Colunas + 1; j++)
                {
                    if(ws.Cells[i,j].Value2 == null)
                    {
                        matriz[i -1, j -1] = "-";
                    }
                    else
                    {
                       
                        matriz[i -1, j - 1] = ws.Cells[i, j].Value2.ToString();
                        matriz[i - 1, j - 1] = matriz[i-1,j-1].Replace("'", " ");
                        matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Replace("|", "/");
                        int quebra = matriz[i - 1, j - 1].IndexOf("\n");
                        int quebrar = matriz[i - 1, j - 1].IndexOf("\r");
                        if (quebra > 0)
                        {
                           //MessageBox.Show("Entrou!!");
                           matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Remove(quebra);
                        }
                        if(quebrar > 0)
                        {
                            //MessageBox.Show("Entrou R!!");
                            matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Remove(quebrar);
                        }
                        

                        /*if(matriz[i - 1, j - 1] == "18902") 
                        {
                            matriz[i - 1, j - 1] = outubro;
                        }
                        if (matriz[i - 1, j - 1] == "33147")
                        {
                            matriz[i - 1, j - 1] = outubro2;
                        }*/
                        int found = 0;
                        /*DateTime.FromOADate
                        if (DateTime.FromOADate(matriz[i-1, j-1].))
                        {
                            DateTime.Parse(matriz[i - 1, j - 1]);
                        }
                         date
                        if (Date.TryParse(matriz[i - 1, j - 1], out DateTime Data))
                        {
                            DateTime.Parse(matriz[i - 1, j - 1]);
                        }*/

                        if(Datas.Contains(j)) //Verifica se a coluna atual foi declarada com data
                        {
                            bool sucess = double.TryParse(matriz[i - 1, j - 1], out double d);
                            if (sucess)
                            {

                                //d = double.Parse(matriz[i - 1, j - 1]);
                                try
                                {
                                    DateTime conv = DateTime.FromOADate(d);
                                    String date = conv.ToString("dd/MM/yyy");
                                    //var date = conv.Date;
                                    matriz[i - 1, j - 1] = date;
                                    //MessageBox.Show(matriz[i - 1, j - 1]);
                                }
                                catch
                                {
                                    matriz[i - 1, j - 1] = "-";
                                }

                            }
                        }

                        //MessageBox.Show(matriz[i - 1, j - 1]);

                   

                        double flutuante;
                        if (double.TryParse(matriz[i - 1, j - 1], out flutuante))
                        {
                            double mynum = double.Parse(matriz[i-1,j-1]);
                            
                            matriz[i - 1, j - 1] = mynum.ToString("0.#############");
                            // = mynum.ToString();
                            matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Replace(".", ",");
                            //matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Trim(charsToTrim);
                        }
                    }
                    int contador = 0;
                    int inicio = 0;
                    for(int y = 1; y <= matriz[i - 1, j - 1].Length;y++)
                    {
                        if (matriz[i - 1, j - 1][y - 1] == ' ' && matriz[i - 1, j - 1][y] == ' ') // Conta o numero de espaços em branco seguidos
                       {
                            contador++;
                            if(contador == 1)
                            {
                                inicio = y - 1;
                            }
                            
                       }
                        else
                        {
                            if(inicio > 0) // caso haja mais d eum espaço em branco 
                            {
                                matriz[i - 1, j - 1] = matriz[i - 1, j - 1].Remove(inicio, contador);
                                contador = 0;
                                inicio = 0;
                            }
                            
                        }
                    }   
                }
                              
            }

            int[] MaiorPalavra = new int[Colunas];
            for(int i = 0; i < Colunas; i++)
            {
                MaiorPalavra[i] = 0;
            }
            
            

            for (int i = 1; i < Linhas; i++)
            { 
                for(int j = 0; j < Colunas; j++)
                {
                    if(matriz[i, j].Length > MaiorPalavra[j])
                    {
                        MaiorPalavra[j] = matriz[i, j].Length; ;
                    }
                    //MessageBox.Show(matriz[i,j]);
                }
            }

            for (int i = 0; i < Colunas; i++)
            {
                //MessageBox.Show(MaiorPalavra[i].ToString());
            }

            //string path = "texto.txt";
            //string path = "C:/Users/Thalles/Desktop/teste/teste2.txt";
            string saida_2 = "C:/Users/Thalles/Desktop/teste/";
            string VerPadrao = Path.GetFileNameWithoutExtension(Padrao);
            string saida = Path.GetFileNameWithoutExtension(path);
            saida_2 = saida_2 + saida + ".txt";
            PastaSaida = PastaSaida + "/" + saida + ".txt";

            if(VerPadrao == "ManterPadrao")
            {
                string vazia = "x                ";
                //MessageBox.Show("Manteve o Padrao!!");

                using (StreamWriter writer = new StreamWriter(PastaSaida))
                {
                    writer.WriteLine("");
                    for (int i = 1; i < Linhas; i++)
                    {
                        for (int j = 0; j < Colunas; j++)
                        {

                            if (j < Colunas - 1)
                            {
                                
                                    int tamanho = matriz[i,j].Length;
                                    int espacos = MaiorPalavra[j] - tamanho;
                                    string branco = new string(' ', espacos + 16);
                                    matriz[i,j] += branco;
                                    writer.Write(matriz[i,j]);
                                
                               
                            }


                            if (j == Colunas - 1)
                            {
                                    writer.Write(matriz[i,j]);  
                            }

                        }
                        writer.WriteLine("");
                    }
                }
                if(ConversaoEmMassa == false)
                {
                    MessageBox.Show("Conversão Realizada com sucesso!");
                }
                return true;
            }
            else
            {
                try
                {
                    // Create an instance of StreamReader to read from a file.
                    // The using statement also closes the StreamReader.
                    using (StreamReader sr = new StreamReader(Padrao))  //"C:/GitHub/Codimuc/Conversor/Conversor/obj/Debug/PadraoA.txt"
                    {
                        string line;
                        // Read and display lines from the file until the end of 
                        // the file is reached
                        int contadorLinhas = 0;
                        while ((line = sr.ReadLine()) != null)
                        {
                            //MessageBox.Show(line);
                            if (contadorLinhas == 1)
                            {
                                int found = line.IndexOf(": ");
                                try
                                {
                                    
                                    MatrizResultado = new string[Linhas, int.Parse(line.Substring(found + 2))];
                                    Atribuicoes = new int[int.Parse(line.Substring(found + 2))];
                                }
                                catch (Exception)
                                {
                                    int ColunaErrada = contadorLinhas - 1;
                                    MessageBox.Show("O valor contido na Coluna " + ColunaErrada + " Não é um número inteiro");
                                    return false;
                                }

                                //MatrizResultado = new string[Linhas, int.Parse(line)];
                                
                                //Atribuicoes = new int[int.Parse(line)];
                                //MessageBox.Show("Tamanho Vetor: " + Atribuicoes.Length.ToString());
                                if(Atribuicoes.Length > Colunas)
                                {
                                    MessageBox.Show("Número de Colunas do Padrão é maior do que o número de Colunas do Arquivo Selecionado para Conversão!!");
                                    return false;
                                }
                            }
                            if (contadorLinhas > 1)
                            {
                                if(contadorLinhas -2 <= Atribuicoes.Length)
                                {
                                    int found = line.IndexOf(": ");
                                    try
                                    {
                                       if (String.IsNullOrEmpty(line.Substring(found + 2))){
                                            Atribuicoes[contadorLinhas - 2] = 0;
                                        }
                                        else
                                        {
                                            Atribuicoes[contadorLinhas - 2] = int.Parse(line.Substring(found + 2));
                                        }
                                        
                                    }
                                    catch (Exception)
                                    {
                                        MessageBox.Show("O valor contido na Coluna " + (contadorLinhas - 1) + " Não é um número inteiro");
                                        return false;
                                    }
                                    
                                    if(Atribuicoes[contadorLinhas -2] > Colunas || Atribuicoes[contadorLinhas - 2] < 0)
                                    {
                                        MessageBox.Show("A Coluna " + (contadorLinhas -1) + "escolhida nao existe no Arquivo selecionado para Conversão!!!");
                                        
                                        return false;
                                    } 
                                    //Atribuicoes[contadorLinhas - 1] = int.Parse(line);
                                    //Atribuicoes[contadorLinhas - 1]--;
                                }

                            }
                            contadorLinhas++;
                            //MessageBox.Show(line);
                        }
                        sr.Close();
                    }
                }
                catch (Exception e)
                {
                    // Let the user know what went wrong.
                    MessageBox.Show("The file could not be read:");
                    MessageBox.Show(e.Message);
                    return false;
                }

                /*for(int i = 0; i < Atribuicoes.Length; i++)
                {
                    MessageBox.Show("Atribuicoes " + i + " = "  + Atribuicoes[i].ToString());
                }*/

                //MessageBox.Show("Chegou aqui!!");
                //MessageBox.Show(Atribuicoes.Length.ToString());
                string vazia = "x                ";
                string finalvazia = "x";

                using (StreamWriter writer = new StreamWriter(PastaSaida))
                {
                    writer.WriteLine("");
                    for (int i = 1; i < Linhas; i++)
                    {
                        for (int j = 0; j < Atribuicoes.Length; j++)
                        {
                            if (j < Atribuicoes.Length - 1)
                            {
                                if (Atribuicoes[j] > 0) //caso exista a coluna do arquivo xls
                                {
                                    int tamanho = matriz[i, Atribuicoes[j] - 1].Length;
                                    int espacos = MaiorPalavra[Atribuicoes[j] - 1] - tamanho;
                                    string branco = new string('-', espacos + 16);
                                    matriz[i, Atribuicoes[j] - 1] += branco;
                                    writer.Write(matriz[i, Atribuicoes[j] - 1]);
                                }
                                else
                                {
                                    writer.Write(vazia);
                                }
                            }


                            if (j == Atribuicoes.Length - 1)
                            {
                                if (Atribuicoes[j] == 0)
                                {
                                    writer.Write(finalvazia);
                                }
                                else
                                {
                                    writer.Write(matriz[i, Atribuicoes[j] - 1]);
                                }
                            }

                        }
                        writer.WriteLine("");
                    }
                    writer.Close();
                }
                if (ConversaoEmMassa == false)
                {
                    MessageBox.Show("Conversão Realizada com sucesso!");
                }
                return true;
            }   
        }
        public void Close()
        {
            wb.Close(0);
        }
    }
}
