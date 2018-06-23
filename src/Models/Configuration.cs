using System;
using System.Collections.Generic;
using System.Text;

namespace CMGenerator.Models
{
    public class Configuration
    {
        public string WorksheetName { get; set; }
        public string ColumnNumber { get; set; }

        public string ColumnResposibleArea { get; set; }

        public string ColumnAction { get; set; }

        public string ColumnPrevisionDate { get; set; }

        public string ColumnConclusionDate { get; set; }

        public string DateFormat { get; internal set; }

        public int PositionNumber { get; internal set; }

        public int PositionResponsibleArea { get; internal set; }

        public int PositionAction { get; internal set; }

        public int PositionPrevisionDate { get; internal set; }

        public int PositionConclusionDate { get; internal set; }

        public static Configuration GetDefault()
        {
            return new Configuration
            {
                WorksheetName = "Controle das Ações",
                ColumnNumber = "CM",
                ColumnResposibleArea = "Área Responsável",
                ColumnAction = "Plano de Ação",
                ColumnPrevisionDate = "Previsão para Conclusão",
                ColumnConclusionDate = "Data da Recebimento e Baixa da Evidência",
                DateFormat = "d"
            };
        }

        internal void CleanPosition()
        {
            PositionNumber = PositionConclusionDate = PositionResponsibleArea = PositionAction = PositionPrevisionDate
                = int.MinValue;
        }

        internal void ValidatedPosition()
        {
            if (PositionNumber == int.MinValue)
                throw new Exception("Informe posição do número CM");

            if (PositionResponsibleArea == int.MinValue)
                throw new Exception("Informe posição da Area Responsavel");

            if (PositionPrevisionDate == int.MinValue)
                throw new Exception("Informe posição da Data de Previsão");

            if (PositionConclusionDate == int.MinValue)
                throw new Exception("Informe posição da Data de Conclusão");
        }
    }
}
