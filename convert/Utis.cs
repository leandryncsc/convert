using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public static void EscreverArquivo(string dados, string nome)
        {
            const string caminho = "./saida";
            const string ext = ".txt";

            CriarPasta(caminho);

            string final = caminho + "/" + nome + ext;

            try
            {
                FileStream fs = File.Create(final);
                fs.Close();

                File.WriteAllText(final, dados);
                Console.WriteLine("\n" + "Gravado " + nome + ext);
                Console.WriteLine("Arquivo escrito com sucesso.");
            }
            catch(Exception ex)
            {
                Console.WriteLine("\n" + "! Erro !\nProblema ao escrever arquivo.");
                Console.WriteLine("\n" + ex.Message);
                //throw ex;
            }
        }

        public static void CriarPasta (string dir)
        {
            if(Directory.Exists(dir))
            {
                return;
            }

            Directory.CreateDirectory(dir);
        }
    }
}
