using CMGenerator.Models;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CMGenerator.Helper
{
    public class ConfigurationFactory
    {
        static Configuration Default = GetDefault();

        static Configuration LessThan2020 = GetLessThan2020();

        public static Configuration Get(FileInfo fileInfo)
        {
            try
            {
                if (fileInfo != null)
                {
                    var resultString = Regex.Match(fileInfo.Name, @"\d+").Value;
                    int year = int.Parse(resultString);
                    if (year < 2020)
                        return LessThan2020;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }

            return Default;
        }

        static Configuration GetDefault()
        {
            return new Configuration
            {
                WorksheetName = "CM",
                WorksheetProductName = "CM",
                WorksheetFsmProductName = "FSM",
                ColumnNumber = "CM",
                ColumnResposibleArea = "Área Responsável",
                ColumnAction = "Plano de Ação",
                ColumnPrevisionDate = "Prazo da Ação",
                ColumnConclusionDate = "Data da Baixa no Onbase",
                ColumnExtensionOne = "1º Prorrogação",
                ColumnExtensionTwo = "2º Prorrogação",
                ColumnExtensionThree = "3º Prorrogação",
                ColumnProduct = "Código",
                ColumnProductDescription = "Descrição",
                ColumnJustification = "Justificativa da Mudança",
                DateFormat = "d",
                RowStart = 6
            };
        }

        static Configuration GetLessThan2020()
        {
            return new Configuration
            {
                WorksheetName = "Controle das Ações",
                WorksheetProductName = "CM",
                WorksheetFsmProductName = "FSM",
                ColumnNumber = "CM",
                ColumnResposibleArea = "Área Responsável",
                ColumnAction = "Plano de Ação",
                ColumnPrevisionDate = "Previsão para Conclusão",
                ColumnConclusionDate = "Data de Conclusão da Ação",
                ColumnExtensionOne = "1º Prorrogação",
                ColumnExtensionTwo = "2º Prorrogação",
                ColumnExtensionThree = "3º Prorrogação",
                ColumnProduct = "Código Descrição",
                ColumnProductDescription = "zzzzzzzzzzzzzzz",
                ColumnJustification = "Justificativa da Mudança",
                DateFormat = "d",
                RowStart = 1
            };
        }
    }
}
