using Microsoft.Office.Interop.Excel;
using ProjetoRe.Exceptions;
using ProjetoRe.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace ProjetoRe.Apps
{
    public class RelatorioValvula
    {
        static string colunaNomeArquivoRelatorio;
        static string formatoNomeArquivoRelatorio;
        static int alturaImagem;
        static int larguraImagem;
        static List<string> meses;

        static RelatorioValvula()
        {
            colunaNomeArquivoRelatorio = ConfigurationManager.AppSettings["ColunaNomeArquivoRelatorio"];
            formatoNomeArquivoRelatorio = ConfigurationManager.AppSettings["FormatoNomeArquivoRelatorio"];
            alturaImagem = Convert.ToInt32(ConfigurationManager.AppSettings["AlturaImagem"]);
            larguraImagem = Convert.ToInt32(ConfigurationManager.AppSettings["LarguraImagem"]);
            meses = new List<string>() {
                "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
            };
        }

        public static void GerarRelatorioDeValvulas(List<MapItem> mapeamentos, List<Dictionary<string, string>> valvulas)
        {
            try
            {
                Directory.CreateDirectory(Configs.UrlDiretorioDestino);
            }
            catch (Exception ex)
            {
                if (!(ex is CustomException))
                    throw new CustomException(String.Format("O diretório {0} já existe. Por favor, exclua o mesmo e tente novamente", Configs.UrlDiretorioDestino));
                else throw ex;
            }

            foreach (var valvula in valvulas)
            {
                GerarRelatorioDeValvula(mapeamentos, valvula);
                Console.WriteLine(String.Format("{0} relatórios gerados", valvulas.IndexOf(valvula) + 1));
            }
        }

        private static void GerarRelatorioDeValvula(List<MapItem> mapeamentos, Dictionary<string, string> valvula)
        {
            Application app = new Application();
            try
            {
                Workbook wb = app.Workbooks.Open(Configs.UrlArquivoTemplate, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                Worksheet sheet = (Worksheet)wb.Sheets[1];

                string propriedadeNomeArquivo = valvula[colunaNomeArquivoRelatorio];
                string nomeArquivoDestino = String.Format(formatoNomeArquivoRelatorio, propriedadeNomeArquivo);
                string arquivoDestino = String.Format("{0}/{1}", Configs.UrlDiretorioDestino, nomeArquivoDestino);

                foreach (MapItem mapeamento in mapeamentos)
                {
                    string valorPropriedade = valvula[mapeamento.PropriedadeOgirem];
                    escreverPropriedade(sheet, mapeamento, valorPropriedade, Configs.UrlDiretorioImagens);
                }

                wb.SaveAs(arquivoDestino, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing,
                    XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                if (!(ex is CustomException))
                    throw new CustomException(String.Format("Houve um erro ao gerar o relatório da válvula {0}.. \n\n {1}", valvula["ID"], ex.Message));
                else throw ex;
            }
            finally
            {
                ExcelAppHelper.KillExcel(app);
            }
        }

        private static void escreverPropriedade(Worksheet sheet, MapItem mapeamento, string valorPropriedade, string diretorioImagens)
        {
            var match = Regex.Match(mapeamento.CelulaDestino, @"(?<linha>\d+)(?<coluna>.+)");
            int linha = Convert.ToInt32(match.Groups["linha"].Value);
            string coluna = match.Groups["coluna"].Value;

            if (mapeamento.Tipo == "Imagem")
            {
                if (valorPropriedade != null && !String.IsNullOrEmpty(valorPropriedade.Trim()))
                {
                    string imagem = String.Format("{0}/{1}", diretorioImagens, valorPropriedade.Trim());
                    adicionarImagem(sheet, imagem, linha, coluna);
                }
            }
            else if (mapeamento.Tipo == "Data_MES_EXTENSO")
            {
                string mesExtenso = recuperarMesPorExtenso(valorPropriedade);
                sheet.Cells[linha, coluna].Value = mesExtenso;
            }
            else if (mapeamento.Tipo == "Data_DIA")
            {
                string dia = recuperarDia(valorPropriedade);
                sheet.Cells[linha, coluna].Value = dia;
            }
            else if (mapeamento.Tipo == "Data_ANO")
            {
                string ano = recuperarAno(valorPropriedade);
                sheet.Cells[linha, coluna].Value = ano;
            }
            else if (mapeamento.Tipo == "Template")
            {
                string texto = String.Format(mapeamento.Template, valorPropriedade);
                sheet.Cells[linha, coluna].Value = texto;
            }
            else
            {
                sheet.Cells[linha, coluna].Value = valorPropriedade;
            }
        }

        private static string recuperarAno(string valorPropriedade)
        {
            return valorPropriedade.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries)[2].Substring(0, 4);
        }

        private static string recuperarDia(string valorPropriedade)
        {
            return Convert.ToInt32(valorPropriedade.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries)[0]).ToString();
        }

        private static string recuperarMesPorExtenso(string valorPropriedade)
        {
            int mes = Convert.ToInt32(valorPropriedade.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries)[1]);
            return meses[mes - 1];
        }

        private static void adicionarImagem(Worksheet sheet, string imagem, int linha, string coluna)
        {
            try
            {
                float left = (float)sheet.Cells[linha, coluna].Left;
                float top = (float)sheet.Cells[linha, coluna].Top;
                Shape shape = sheet.Shapes.AddPicture(imagem, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, larguraImagem, alturaImagem);
                shape.Locked = false;
            }
            catch (Exception ex)
            {
                throw new CustomException(String.Format("A imagem '{0}' não foi encontrada", imagem));
            }
        }
    }
}
