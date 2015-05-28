using InDesign;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using indd = InDesign;

namespace Teste_Indesign
{
    public partial class Form1 : Form
    {
       public static indd.ApplicationClass app;
       public static List<string> erros = new List<string>();
        public Form1()
       {
           try
           {
               new indd.Application();
           }
           catch { 
           }
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                

                app = new indd.ApplicationClass();
                app.ScriptPreferences.UserInteractionLevel = idUserInteractionLevels.idNeverInteract; //Desligar interação do usuário
                
            }
            catch {
                MessageBox.Show("Por favor, abra o Indesign");
                Process.GetCurrentProcess().Kill();
            }

        }

        public void exportarCurso(String caminho){
            List<String> pastadeAlbuns = new List<String>();
            List<String> arquivosIndd = new List<String>();
            if (Directory.Exists(caminho)) {  // Verifcando se o caminho inserido é válido
                pastadeAlbuns.AddRange(Directory.GetDirectories(caminho, "20*"));
                foreach (String s in pastadeAlbuns){
                    arquivosIndd.AddRange(Directory.GetFiles(s,"*.indd"));
                    }
                progressBar1.Maximum = arquivosIndd.Count;
                foreach (String a in arquivosIndd){
                    Document document = (Document) app.Open(a);
                    exportar(document,caminho);
                    progressBar1.Value++;
                    }
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(textBox1.Text + "//Erros ao exportar.txt"))
            {
                foreach (String e in erros)
                {
                    file.WriteLine(e);
                }
            }
        }


       
        public void exportar(Document doc,String caminho) {
            String temp = "";
            String parentFolder = doc.FilePath;
            String[] parcial = parentFolder.Split('\\'); // Obtendo o nome do album, extraindo com caminho geral do arquivo
            String nameAlbum = parcial[parcial.Length - 1];
            label1.Text = "Álbum atual: "+nameAlbum;
            //Atualizar todos os links
        
            for (int i = 1; i <= doc.Links.Count; i++ ){
                Link link = (Link)doc.Links[i];
                
                try {
                    link.Update();
                }
                catch {
                    
                    erros.Add("Link quebrado no álbum "+ nameAlbum+" "+ link.FilePath);
                }
            }

            String destinoPasta = "";
            destinoPasta = caminho + "\\TRATADO\\" + nameAlbum;
            if (!Directory.Exists(destinoPasta))
            {
                Directory.CreateDirectory(destinoPasta);
            }

            for (int cont = 1; cont <= doc.Pages.Count; cont++)
            {
                Page page = (Page)doc.Pages[cont];
                String myPageName = page.Name;
                
                String destino = "";
                destino = caminho + "\\TRABALHADO\\" + nameAlbum+"\\";

                if (!Directory.Exists(destino))
                {
                    Directory.CreateDirectory(destino);
                }


                app.JPEGExportPreferences.PageString = myPageName;
                app.JPEGExportPreferences.JPEGQuality = idJPEGOptionsQuality.idMaximum;
                app.JPEGExportPreferences.ExportResolution = 300;
                app.JPEGExportPreferences.JPEGExportRange = idExportRangeOrAllPages.idExportRange;
                

                if (cont < 100)
                {
                    myPageName = "0" + myPageName;
                }
                if (cont < 10)
                {
                    myPageName = "0" + myPageName;
                }
                temp = doc.Name.Substring(0, doc.Name.Length - 5) + " " + myPageName;
                myPageName = temp;
                temp = doc.Name.Substring(0, doc.Name.Length - 5);


                doc.Export(idExportFormat.idJPG, destino + myPageName + ".jpg");
               
            }
            app.ActiveDocument.Close(idSaveOptions.idNo);
          
            moverArquivos(parentFolder,destinoPasta);
        }

        private void moverArquivos(String orign, String target){
            //String handle = "";
            List<String> arquivosPasta = new List<string>();
            if(Directory.Exists(orign)){
                arquivosPasta.AddRange(Directory.GetFiles(orign,"*"));
            }
            if (!Directory.Exists(target))
            {
                Directory.CreateDirectory(target);
            }

            foreach(String s in arquivosPasta){

                String[] parcial = s.Split('\\'); // Obtendo o nome do album, extraindo com caminho geral do arquivo
                String fileName = parcial[parcial.Length - 1];
                try
                {
                    File.Move(s, target + "\\" + fileName);
                }
                catch { 

                }
            }

           // Directory.Delete(orign);

        }

        private InDesign.Application COMCreateObject(string p)
        {
            throw new NotImplementedException();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            exportarCurso(textBox1.Text);
            label1.Text = "Insira o caminho para o contrato";
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (Directory.Exists(textBox1.Text)) { label1.Text = "Caminho Válido"; }
            else { label1.Text = "Caminho inválido"; }
        }
    }
}
