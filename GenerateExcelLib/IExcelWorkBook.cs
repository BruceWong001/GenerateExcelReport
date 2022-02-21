using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GenerateExcelLib
{

    public class DrawParameter
    {
        public int StartRow{get;set;}
        public int StartCol{get;set;}
        public Dictionary<string, MergeCell> MergeCells{get;set;}
        public List<int> HiddenColumns{get;set;}
        
    }

    public class MergeCell : IComparable<MergeCell>
    {
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }

        public MergeCell(int startRow, int startColumn, int totalRows, int totalColumns)
        {
            if (startRow < 0 || startColumn < 0 || totalRows < 0 || totalColumns < 0)
            {
                throw new ArgumentOutOfRangeException("These is negative paramete.");
            }
            this.StartRow = startRow;
            this.StartColumn = startColumn;
            this.TotalRows = totalRows;
            this.TotalColumns = totalColumns;
        }

        public void AddOffSet(int offSetRows, int offSetColumns)
        {
            if (this.StartRow + offSetRows < 0 || this.StartColumn + offSetColumns < 0)
            {
                throw new ArgumentOutOfRangeException("The start row or start column is less than zero after added offset.");
            }
            this.StartRow += offSetRows;
            this.StartColumn += offSetColumns;
        }

        public int CompareTo(MergeCell mergeCell)
        {
            if (this.StartRow == mergeCell.StartRow &&
                this.StartColumn == mergeCell.StartColumn &&
                this.TotalRows == mergeCell.TotalRows &&
                this.TotalColumns == mergeCell.TotalColumns)
            {
                return 0;
            }
            return 1;
        }
    }

    ///
    /// this interface is for extracting the common behaivors which is for drawing a phyasical excel file.
    ///
    public interface IExcelWorkBook
    {
        ///auto append data following the latest row.
        /// start row and col are one.
        void DrawExcel(DataTable data,Boolean hasHeader=true);
        /// specify the location to draw data in excel, it is used when the first data is drawed.
        /// start Row and Col are based on one.
        void DrawExcel(DataTable data,int startRow,int startCol,Boolean hasHeader=true);
        /// specify the parameter obj to draw in the excel,it is used when the first data is drawed.
        /// the parameter will be applyed in the rest call, so no need to call this function everytime.
        void DrawExcel(DataTable data,DrawParameter parameter ,Boolean hasHeader=true);
        ///must call this action to save all the change in excel.
        void Save();

    }
}