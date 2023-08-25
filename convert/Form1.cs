using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using Bytescout.Spreadsheet;

namespace convert
{
    public partial class Form1 : Form
    {
        private static Button bconverter = null;
        private static Label titulo = null;

        public Form1()
        {
            InitializeComponent();
            DefinirInterface();
            DefinirControles();
        }

        public static void Preparar ()
        {
            string arq = Utis.PegarArquivo();

            if(string.IsNullOrEmpty(arq))
            {
                return;
            }

            Spreadsheet doc = new Spreadsheet();
            doc.LoadFromFile(arq);

            Worksheet ws = doc.Workbook.Worksheets[0];

            string dados = TratarDados(ws);

            /* Se retornar nulo é por conta de um erro ao ler
            os dados da planilha (incompatibilidade, etc.) */
            if(string.IsNullOrEmpty(dados))
            {
                MessageBox.Show("Arquivo incompatível.","Erro catastrófico");
                bconverter.Focus();
                return;
            }

            string sufixo = "_saida";
            string arq_nome = Path.GetFileName(arq.Split('.')[0]);

            Utis.SalvarArquivo(arq_nome + sufixo, dados);

            doc.Close();
        }

        private static string TratarDados (Worksheet ws)
        {
            string cabeca = string.Empty;
            string corpo = string.Empty;

            // Define se deve coletar dados para o cabeçalho
            bool catoblepas = true;

            string tag = string.Empty;
            
            int ylin = 0;

            const int MAX_TENTATIVAS = 300;

            // Conta quantas vezes ler a tag foram tentadas
            int zero_tentativas = 0;
            
            for(int xcol = 0; xcol < MAX_TENTATIVAS; xcol++)
            {
                // Cell(y, x)
                string celula = ws.Cell(ylin, xcol).ToString();
                string _valor = celula;

                Console.WriteLine("COL " + xcol);

                Action pularLinha = () =>
                {
                    ylin++;
                    xcol = -1;
                };

                // Coleta a "tag" da linha
                if(xcol == 0)
                {
                    Console.WriteLine("CELL " + celula);
                    zero_tentativas++;
                    
                    if(zero_tentativas > MAX_TENTATIVAS)
                    {
                        return null;
                    }

                    tag = celula;
                }

                // Se estiver vazio, pula para a próxima
                if(string.IsNullOrEmpty(tag))
                {
                    pularLinha();
                    continue;
                }

                Console.WriteLine(tag + ": " + celula);

                // Pula para a próxima linha
                if(celula == "*")
                {
                    pularLinha();

                    if(catoblepas)
                    {
                        cabeca += _valor;
                        catoblepas = false;
                        continue;
                    }

                    _valor += "\n";
                    Console.Write('\n');

                    // Tag 04 delimita o fim do arquivo
                    if(tag == "04")
                    {
                        corpo += _valor;
                        break;
                    }
                }

                /* Se a regra de inserir no cabeçalho estiver
                desativada, então insere no "corpo". */
                if(!catoblepas)
                {
                    corpo += _valor;
                    continue;
                }

                cabeca += _valor;
            }

            Console.WriteLine(string.Empty);
            Console.WriteLine("CABEÇALHO: " + cabeca + "\n");
            Console.WriteLine("CORPO: " + corpo);

            return cabeca + "\n" + corpo;
        }

        /* CONTROLES E UI */

        private void DefinirInterface ()
        {
            Color fazul = Color.FromArgb(0, 86, 163);
            Color fvermelho = Color.FromArgb(216, 11, 19);

            this.BackColor = Color.White;

            int X = this.Width;
            int Y = this.Height;

            const int bX = 160;
            const int bY = 70;

            bconverter = new Button()
            {
                Name = "bconverte",
                Text = "Converter",
                Width = bX,
                Height = bY,
                Location = new Point(X / 2 - bX / 2, (Y / 2 - bY / 2) + 25),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                ForeColor = fazul,
                //BackColor = fvermelho,
            };

            const int lX = 500;
            const int lY = 100;

            titulo = new Label()
            {
                Text = "Converte XLS para TXT",
                Width = lX,
                Height = lY,
                Font = new Font(this.Font.Name, 20, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(X/2 - lX/2, bconverter.Width - 86),
                ForeColor = fazul,
            };

            this.Controls.Add(bconverter);
            this.Controls.Add(titulo);
        }

        private void DefinirControles ()
        {
            // Ao fechar o formulário
            this.FormClosing += (s, e) =>
            {
                titulo.Font.Dispose();
            };

            bconverter.Click += (s, e) =>
            {
                Preparar();
            };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Preparar();
        }
    }
}
