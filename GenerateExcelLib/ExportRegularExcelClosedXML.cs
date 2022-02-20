using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Threading.Tasks;

namespace GenerateExcelLib
{
    public class ExportRegularExcelClosedXML : IExcelWorkBook
    {
        private Stream m_IOStream;
        private XLWorkbook m_workBook;
        private int currentRow_Offset=1;
        private int currentCol_Start=1;
        private DrawParameter m_DrawParameter;
        public ExportRegularExcelClosedXML(Stream stream)
        {
                m_IOStream=stream;
                m_workBook= new XLWorkbook();           

            
        }
        ~ExportRegularExcelClosedXML()
        {
            m_workBook?.Dispose();
        }
        public void DrawExcel(DataTable data, bool hasHeader = true)
        {
            DrawExcel(data,currentRow_Offset,currentCol_Start,hasHeader);     
            
        }

        public void DrawExcel(DataTable data, int startRow, int startCol, bool hasHeader = true)
        {
            var worksheet = m_workBook.Worksheets.Add("Sheet1");
            int ColCount = data.Columns.Count; 
            int RowCount = data.Rows.Count;
            int StartRowNum=startRow;
            int StartColNum=startCol;
            // draw the header
            if(hasHeader)
            {
                for (int i = 0; i < ColCount; i++)
                {
                    worksheet.Cell(StartRowNum, StartColNum+i).Value=data.Columns[i].ColumnName; //

                }
            }
            int rowOffset=hasHeader?1:0;
            // draw the data body
            for (int rowIndex = 0; rowIndex < RowCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < ColCount; colIndex++)
                {
                    worksheet.Cell(StartRowNum+rowIndex+rowOffset, StartColNum+colIndex).Value =data.Rows[rowIndex][colIndex].ToString(); //
                }
            }
            worksheet.Rows().AdjustToContents(); //
            // merge cell
            // if(m_DrawParameter!=null)
            // {   
            //     if (m_DrawParameter.MergeCells!=null)
            //     {
            //         foreach(var mergedCell in m_DrawParameter.MergeCells.Values)
            //         {
            //             worksheet. cells.Merge(StartRowNum+rowOffset+mergedCell.Item2,StartColNum+mergedCell.Item1,
            //                         mergedCell.Item4,mergedCell.Item3);
            //         }
            //     }
            // }

            //calculate the new position of row and column.
            currentRow_Offset+=StartRowNum+rowOffset+RowCount;
            currentCol_Start=startCol;
        }

        public void DrawExcel(DataTable data, DrawParameter parameter, bool hasHeader = true)
        {
            throw new NotImplementedException();
        }
        ///
        /// ClosedXML save uses save as
        ///
        public void Save()
        {
            try
            {
                m_workBook.SaveAs(m_IOStream);
            }
            catch(Exception ex)
            {
                throw new InvalidDataException("The IO Stream is invalidate, so it cannot create a Workbook",ex);
            }
        }
    }
}