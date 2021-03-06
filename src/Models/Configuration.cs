﻿using System;

namespace CMGenerator.Models
{
    public class Configuration
    {
        public string WorksheetName { get; set; }

        public string WorksheetProductName { get; set; }

        public string WorksheetFsmProductName { get; set; }

        public string ColumnNumber { get; set; }

        public string ColumnResposibleArea { get; set; }

        public string ColumnAction { get; set; }

        public string ColumnPrevisionDate { get; set; }

        public string ColumnConclusionDate { get; set; }

        public string ColumnExtensionOne { get; set; }

        public string ColumnExtensionTwo { get; set; }

        public string ColumnExtensionThree { get; set; }

        public string ColumnProduct { get; set; }

        public string ColumnProductDescription { get; set; }

        public string ColumnJustification { get; internal set; }

        public string DateFormat { get; internal set; }

        public int RowStart { get; internal set; }

        public int PositionNumber { get; internal set; }

        public int PositionResponsibleArea { get; internal set; }

        public int PositionAction { get; internal set; }

        public int PositionPrevisionDate { get; internal set; }

        public int PositionConclusionDate { get; internal set; }

        public int PositionExtensionOne { get; internal set; }

        public int PositionExtensionTwo { get; internal set; }

        public int PositionExtensionThree { get; internal set; }

        public int PositionProduct { get; internal set; }

        public int PositionProductDescription { get; internal set; }

        public int PositionJustification { get; internal set; }

        internal void CleanPosition()
        {
            PositionNumber = PositionConclusionDate = PositionResponsibleArea = PositionAction = PositionPrevisionDate
                = PositionExtensionOne = PositionExtensionTwo = PositionExtensionThree = PositionProduct 
                = PositionProductDescription = PositionJustification
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
