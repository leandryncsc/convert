using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Bytescout.Spreadsheet;

namespace convert
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Cria uma nova planilha a partir do arquivo SimpleReport.xls
            string caminho = "C:\\Users\\Administrador\\Desktop\\teste.xlsx";
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(caminho);
            string novoarquivo = "C:\\Users\\Administrador\\Desktop\\jslaima.txt";
            // exclui o arquivo de saída se já existir
            if (File.Exists(novoarquivo))
            {
                File.Delete(novoarquivo);
            }

            // salva em TXT

            document.Workbook.Worksheets[0].SaveAsTXT(novoarquivo);


            Thread.Sleep(5);

            // abre o documento de saída no visualizador padrão
            if (File.Exists(novoarquivo))
            {
                Process.Start(novoarquivo);
            }
        }

  
    }
}
