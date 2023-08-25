using System;
using System.IO;
using System.Windows.Forms;

namespace convert
{
    internal class Utis
    {
        public static string PegarArquivo ()
        {
            using (OpenFileDialog fd = new OpenFileDialog()
            {
                //InitialDirectory = @"C:\",
                RestoreDirectory = true,
                Title = "Abrir planilha",
                Filter = "Planilha (*.xls) (*.xlsx)|*.xls;*.xlsx",
            })
            {
                DialogResult res = fd.ShowDialog();

                // Sem arquivo
                if (res == DialogResult.Cancel)
                {
                    Console.WriteLine("Sem arquivo");
                    return string.Empty;
                }

                return fd.FileName;
            }
        }

        public static void SalvarArquivo (string nome, string dados)
        {
            const string ext = ".txt";

            SaveFileDialog sfd = new SaveFileDialog()
            {
                Title = "Salvar arquivo convertido",
                FileName = nome + ext,
                Filter = "TXT (*.txt)|*.txt|TBF (*.tbf)|*.tbf",
                FilterIndex = 0,
                RestoreDirectory = true,
            };

            if(sfd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string dir = sfd.FileName;

            EscreverDados(dir, dados);
        }

        private static void EscreverDados (string dir, string dados)
        {
            try
            {
                FileStream fs = File.Create(dir);
                fs.Close();

                File.WriteAllText(dir, dados);
                Console.WriteLine("\n" + "Gravado: " + dir);
                Console.WriteLine("Arquivo escrito com sucesso.");

                MessageBox.Show("Arquivo salvo.", "Aviso");
            }
            catch(Exception ex)
            {
                Console.WriteLine("\n" + "! Erro !\nProblema ao escrever arquivo.");
                Console.WriteLine("\n" + ex.Message);
            }
        }

        /*public static void CriarPasta (string dir)
        {
            if(Directory.Exists(dir))
            {
                return;
            }

            Directory.CreateDirectory(dir);
        }*/
    }
}
