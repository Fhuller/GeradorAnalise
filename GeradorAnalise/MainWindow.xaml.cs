using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Xceed.Words.NET;

namespace GeradorAnalise
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BTN1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ValidaCampos();
                var arquivo = CriaPastaEArquivo();
                EditaArquivo(arquivo);

                MessageBox.Show("Arquivo criado com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Um erro ocorreu: " + ex.Message, "Exception Sample", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            TBX4.Text = dialog.SelectedPath;
        }

        public string GeraNomeArquivo()
        {
            string novoNome = $@"Analise_Tecnica_{TBX1.Text}_{GeraNomeAnalise(TBX3.Text)}.docx";

            return novoNome;
        }

        public string GeraNomeAnalise(string str)
        {
            string nome = "";

            foreach(string word in str.Split(' '))
            {
                nome += char.ToUpper(word[0]) + word.Substring(1) + "_";
            }

            return nome.TrimEnd('_');
        }

        public string CorrigeTituloTask(string str)
        {
            string nome = "";

            foreach (string word in str.Split(' '))
            {
                nome += char.ToUpper(word[0]) + word.Substring(1) + " ";
            }

            return nome.TrimEnd(' ');
        }

        public void ValidaCampos()
        {
            if (TBX1.Text == "")
            {
                throw new System.Exception("Numero da task vazia!");
            }
            if (TBX2.Text == "")
            {
                throw new System.Exception("Flo vazio!");
            }
            if (TBX3.Text == "")
            {
                throw new System.Exception("Nome da task vazia!");
            }
            if (TBX4.Text == "")
            {
                throw new System.Exception("Caminho do arquivo vazio!");
            }
        }

        public string CriaPastaEArquivo()
        {
            string destinyLocation = $@"{TBX4.Text}\{TBX1.Text}";
            string fileName = GeraNomeArquivo();

            if (Directory.Exists(destinyLocation))
                throw new System.Exception("Arquivo já existente");

            System.IO.Directory.CreateDirectory($"{destinyLocation}");

            string dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            string sourceLocation = $@"{dir}\ArquivoPadrao\PadraoClone.docx";

            string fileLocation = $@"{destinyLocation}\{fileName}";

            System.IO.File.Copy(sourceLocation, fileLocation, true);

            return fileLocation;
        }

        public void EditaArquivo(string caminhoArquivo)
        {
            string caminho = caminhoArquivo.Replace(@"/", @"\");

            using(DocX documento = DocX.Load(caminho))
            {
                documento.ReplaceText("#task", TBX1.Text);
                documento.ReplaceText("#flo", TBX2.Text);
                documento.ReplaceText("#title", CorrigeTituloTask(TBX3.Text));

                documento.SaveAs(caminho);
            }
        }
    }
}
