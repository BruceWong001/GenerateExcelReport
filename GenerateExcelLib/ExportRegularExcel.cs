using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using System.Threading.Tasks;

namespace GenerateExcelLib
{
    public class ExportRegularExcel:IExcelWorkBook
    {
        private Stream m_IOStream;
        public ExportRegularExcel(Stream stream)
        {
            try
            {
                m_IOStream=stream;
                m_workBook= new Workbook(m_IOStream);                
            }
            catch(Exception ex)
            {
                throw new InvalidDataException("The IO Stream is invalidate, so it cannot create a Workbook",ex);
            }
            
        }


        ///
        /// Generate excel from DataTable. data column name should be table header in excel.
        /// if generate from existing template, you can set needHead=false to only write data body.
        ///
        private int currentRow_Offset=1;
        private int currentCol_Start=1;
        private Workbook m_workBook;
        private DrawParameter m_DrawParameter;

        ~ExportRegularExcel()
        {
            m_workBook?.Dispose();
        }
        public void DrawExcel(DataTable data,Boolean hasHeader=true)
        {
            DrawExcel(data,currentRow_Offset,currentCol_Start,hasHeader);            
        }

        public void DrawExcel(DataTable data, int startRow, int startCol, bool hasHeader = true)
        {
            Workbook book = m_workBook; 
            Worksheet sheet = m_workBook.Worksheets[0]; 
            Cells cells = sheet.Cells; 
            int ColCount = data.Columns.Count; 
            int RowCount = data.Rows.Count;
            int StartRowNum=startRow-1;
            int StartColNum=startCol-1;
            // draw the header
            if(hasHeader)
            {
                for (int i = 0; i < ColCount; i++)
                {
                    cells[StartRowNum, StartColNum+i].PutValue(data.Columns[i].ColumnName); //

                }
            }
            int rowOffset=hasHeader?1:0;
            // draw the data body
            for (int rowIndex = 0; rowIndex < RowCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < ColCount; colIndex++)
                {
                    cells[StartRowNum+rowIndex+rowOffset, StartColNum+colIndex].PutValue(data.Rows[rowIndex][colIndex].ToString(),true,true); //
                }
            }
            sheet.AutoFitColumns(); //
            // merge cell
            if(m_DrawParameter!=null)
            {   
                if (m_DrawParameter.MergeCells!=null)
                {
                    foreach(var mergedCell in m_DrawParameter.MergeCells.Values)
                    {
                        cells.Merge(StartRowNum+rowOffset+mergedCell.Item2,StartColNum+mergedCell.Item1,
                                    mergedCell.Item4,mergedCell.Item3);
                    }
                }
            }

            //calculate the new position of row and column.
            currentRow_Offset+=StartRowNum+rowOffset+RowCount;
            currentCol_Start=startCol;
        }

        public void DrawExcel(DataTable data, DrawParameter parameter, bool hasHeader = true)
        {
            m_DrawParameter=parameter;
            DrawExcel(data,currentRow_Offset>1?currentRow_Offset:parameter.StartRow,parameter.StartCol>1?parameter.StartCol:currentCol_Start,hasHeader);   

        }
        public void Save()
        {
            //hide or delete columns before we impletement Save action
            if(m_DrawParameter!=null)
            {
                if(m_DrawParameter.HiddenColumns!=null)
                {
                    Boolean isFirstColumn=true;
                    m_DrawParameter.HiddenColumns.ForEach((columnIndex)=>{
                        m_workBook.Worksheets[0].Cells.DeleteColumn(isFirstColumn?columnIndex:columnIndex-1);
                        isFirstColumn=false;
                    });
                }
            }
            //
            m_workBook.Save(m_IOStream, SaveFormat.Xlsx);
        }

    }
}
