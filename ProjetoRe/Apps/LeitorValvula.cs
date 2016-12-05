using Microsoft.Office.Interop.Excel;
using ProjetoRe.Exceptions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace ProjetoRe.Apps
{
    public class LeitorValvula
    {
        public static List<Dictionary<string, string>> LerValvulas()
        {
            Application app = new Application();
            try
            {
                List<Dictionary<string, string>> valves = new List<Dictionary<string, string>>();
                Workbook wb = app.Workbooks.Open(Configs.UrlArquivoValvulas, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                Worksheet sheet = (Worksheet)wb.Sheets[1];

                int linhaPropriedades = 8;
                List<string> propriedades = lerPropriedades(linhaPropriedades, sheet);

                int inicioLinhaValvulas = linhaPropriedades + 1;
                for (int rowNumber = inicioLinhaValvulas; true; rowNumber++)
                {
                    bool validRow = sheet.Cells[rowNumber, 1].Value != null && !String.IsNullOrEmpty(sheet.Cells[rowNumber, 1].Value.ToString());
                    if (!validRow)
                        break;

                    Dictionary<string, string> valvula = new Dictionary<string, string>();
                    for (int propNumber = 0; propNumber < propriedades.Count; propNumber++)
                    {
                        string propertyValue = sheet.Cells[rowNumber, propNumber + 1].Value == null ? String.Empty : sheet.Cells[rowNumber, propNumber + 1].Value.ToString();
                        valvula.Add(propriedades[propNumber], propertyValue);
                    }

                    valves.Add(valvula);
                }

                wb.Close();
                Console.WriteLine(String.Format("Foram encontradas {0} válvulas.", valves.Count));
                return valves;
            }
            catch (Exception ex)
            {
                throw new CustomException("Houve um erro ao ler as válvulas.. \n\n " + ex.Message);
            }
            finally
            {
                ExcelAppHelper.KillExcel(app);
            }
        }

        private static List<string> lerPropriedades(int linhaPropriedades, Worksheet sheet)
        {
            List<string> propriedades = new List<string>();
            for (int propIndex = 1; true; propIndex++)
            {
                string propName = sheet.Cells[linhaPropriedades, propIndex].Value;
                if (String.IsNullOrEmpty(propName))
                    break;

                propriedades.Add(propName);
            }

            return propriedades;
        }
    }
}
