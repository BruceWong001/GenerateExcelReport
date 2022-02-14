using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace GenerateExcelLib
{
    public class ExportRegularExcel : IDisposable
    {
        private Workbook _book;

        public ExportRegularExcel()
        {
            _book = new Workbook();
        }

        /// <summary>
        /// allow create excel from exist template
        /// </summary>
        /// <param name="stream"></param>
        public ExportRegularExcel(Stream stream)
        {
            _book = new Workbook(stream);
        }

        ///
        /// Generate excel from DataTable. data column name should be table header in excel.
        /// if generate from existing template, you can set needHead=false to only write data body.
        ///
        public Workbook GenerateExcel(DataTable data,int startRow=1,Boolean needHead=true)
        {
            
            Worksheet sheet = _book.Worksheets[0]; 
            Cells cells = sheet.Cells; 


            int Colnum = data.Columns.Count; 
            int Rownum = data.Rows.Count;
            int startRow_position=startRow-1; // by default row should start from the first first row.
            // draw the header
            if(needHead)
            {
                for (int i = 0; i < Colnum; i++)
                {
                    cells[startRow_position, i].PutValue(data.Columns[i].ColumnName); //

                }
                startRow_position++; // if there is header row, the data row should start from second row.
            }
            // draw the data body
            for (int i = 0; i < Rownum; i++)
            {
                for (int k = 0; k < Colnum; k++)
                {
                    cells[startRow_position + i, k].PutValue(data.Rows[i][k].ToString(),true,true); //
                }
            }
            sheet.AutoFitColumns(); //
            return _book; //      

        }
        ///
        /// merge specified cell in excel.
        /// The Workbook should be a avaiable object. all int parameter should be greater than zero, otherwise will throw the exception.
        /// Note:the index from 1, not 0
        ///
        public void MergeCell(Workbook wb,int firstCol,int firstRow,int totalCols, int totalRows)
        {
            if(!((firstCol>0)&&(firstRow>0)&&(totalCols>0)&&(totalRows>0)))
            {
                throw new ArgumentException("the parameters should be greater than zero .");
            }

            wb.Worksheets[0].Cells.Merge(firstRow-1,firstCol-1,totalRows,totalCols);

        }
        public void MergeCell(Workbook wb,Dictionary<string,Tuple<int,int,int,int>> mergedCells,Boolean includedHead=true)
        {
            int StartPosition=includedHead?1:0;
            foreach(var coordinate in mergedCells.Values)
            {
                var startColNum=coordinate.Item1+1;
                var startRowNum=coordinate.Item2+1+StartPosition;
                
                MergeCell(wb,startColNum,startRowNum,coordinate.Item3,coordinate.Item4);
            }
        }

        public void Dispose()
        {
            _book?.Dispose();
        }
    }
}
