using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Resources;
using System.Configuration;
using ProjetoRe.Model;
using System.Web.Script.Serialization;
using ProjetoRe.Apps;
using ProjetoRe.Exceptions;
using static System.Net.WebRequestMethods;
using System.Windows.Forms;

namespace ProjetoRe
{
    class Program
    {
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            try
            {
                configurarDiretorioArquivoValvulas();
                configurarDiretorioArquivoTemplate();
                configurarDiretorioImagens();

                string mapeamentosStr = ConfigurationManager.AppSettings["Mapeamentos"];
                List<MapItem> mapeamentos = new JavaScriptSerializer().Deserialize<List<MapItem>>(mapeamentosStr);

                List<Dictionary<string, string>> valvulas = LeitorValvula.LerValvulas();
                validateMappings(mapeamentos, valvulas);
                RelatorioValvula.GerarRelatorioDeValvulas(mapeamentos, valvulas);
            }
            catch (CustomException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Mensagem);
            }

            Console.WriteLine("FINALIZADO");
            Console.ReadKey();
        }

        private static void configurarDiretorioImagens()
        {
            Console.WriteLine("Selecione o diretório de imagens");
            System.Threading.Thread.Sleep(1500);

            FolderBrowserDialog directorySelectPopUp = new FolderBrowserDialog();
            directorySelectPopUp.SelectedPath = Configs.UrlDiretorioRaiz;
            if (directorySelectPopUp.ShowDialog() == DialogResult.OK)
            {
                Configs.UrlDiretorioImagens = directorySelectPopUp.SelectedPath;
            }
            else
                throw new CustomException("Nenhum diretório de imagens foi escolhido!!");
        }

        private static void configurarDiretorioArquivoTemplate()
        {
            Console.WriteLine("Selecione o arquivo de template");
            System.Threading.Thread.Sleep(1500);

            OpenFileDialog fileSelectPopUp = new OpenFileDialog();
            fileSelectPopUp.Title = "";
            fileSelectPopUp.InitialDirectory = Configs.UrlDiretorioRaiz;
            fileSelectPopUp.Filter = "All EXCEL FILES (*.xls*)|*.xls*";
            fileSelectPopUp.FilterIndex = 1;
            fileSelectPopUp.RestoreDirectory = true;

            if (fileSelectPopUp.ShowDialog() == DialogResult.OK)
            {
                Configs.UrlArquivoTemplate = fileSelectPopUp.FileName;
            }
            else
                throw new CustomException("Nenhum arquivo de template foi escolhido!!");
        }

        private static void configurarDiretorioArquivoValvulas()
        {
            Console.WriteLine("Selecione o arquivo de válvulas");
            System.Threading.Thread.Sleep(1500);

            OpenFileDialog fileSelectPopUp = new OpenFileDialog();
            fileSelectPopUp.Title = "";
            //fileSelectPopUp.InitialDirectory = diretorioRaiz;
            fileSelectPopUp.Filter = "All EXCEL FILES (*.xls*)|*.xls*";
            fileSelectPopUp.FilterIndex = 1;
            fileSelectPopUp.RestoreDirectory = true;

            if (fileSelectPopUp.ShowDialog() == DialogResult.OK)
            {
                Configs.UrlArquivoValvulas = fileSelectPopUp.FileName;
                string nomeDoArquivo = Configs.UrlArquivoValvulas.Split(new string[] { "/", "\\" }, StringSplitOptions.RemoveEmptyEntries).Last();
                int indiceNomeArquivo = Configs.UrlArquivoValvulas.IndexOf(nomeDoArquivo);
                Configs.UrlDiretorioRaiz = Configs.UrlArquivoValvulas.Substring(0, indiceNomeArquivo - 1);
                Configs.UrlDiretorioDestino = String.Format("{0}/{1}", Configs.UrlDiretorioRaiz, nomeDoArquivo.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries)[0]);
            }
            else
                throw new CustomException("Nenhum arquivo de válvulas foi escolhido!!");
        }

        private static void validateMappings(List<MapItem> mapeamentos, List<Dictionary<string, string>> valvulas)
        {
            if (!mapeamentos.Any())
            {
                throw new CustomException("Nenhum mapeamento foi configurado");
            }

            IEnumerable<string> mapeamentosInvalidos = mapeamentos.Where(mapeamento => !valvulas.First().ContainsKey(mapeamento.PropriedadeOgirem)).Select(mapeamento => mapeamento.PropriedadeOgirem);
            if (mapeamentosInvalidos.Any())
            {
                throw new CustomException(String.Format("Existem mapeamentos inválidos: '{0}'", String.Join("', '", mapeamentosInvalidos)));
            }
        }

        //[STAThreadAttribute]
        //static void Main(string[] args)
        //{
        //    string file = @"C:\Users\Felipe\Desktop\id_25642874.xls";
        //    string fileToSave = @"C:\Users\Felipe\Desktop\testes\id_25642874_20.xls";





        //    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing);

        //    Worksheet sheet = (Worksheet)wb.Sheets[1];
        //    sheet.Cells[3, "D"] = "AV DE TESTE";
        //    //sheet.Shapes.AddPicture(@"C:\Users\Felipe\Desktop\Arquivos_Re\01-11-16~files\2016-11-1--8-52-11_A.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 185, 42);
        //    string imageUrl = @"C:\Users\Felipe\Desktop\Arquivos_Re\01-11-16~files\2016-11-1--8-52-11_A.jpg";
        //    //ExcelLibrary.SpreadSheet.Image img1 = ExcelLibrary.SpreadSheet.Image.FromFile(imageUrl);
        //    //System.Windows.Forms.Clipboard.SetDataObject(img1, true);
        //    //sheet.Paste(sheet.Cells[19, "A"], imageUrl);
        //    float left = (float)sheet.Cells[19, "A"].Left;
        //    float top = (float)sheet.Cells[19, "A"].Top;
        //    float width = 477;//768
        //    float height = 681;//1280
        //    Shape shape = sheet.Shapes.AddPicture(imageUrl, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);
        //    shape.Locked = false;

        //    //foreach (Range row in sheet.Rows)
        //    //{
        //    //    String[] rowData = new String[row.Columns.Count];
        //    //    for (int i = 0; i < row.Columns.Count; i++)
        //    //        rowData[i] = row.Cells[1, i + 1].Value2.ToString();
        //    //}

        //    wb.SaveAs(fileToSave, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing,
        //    XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //}
    }
}
