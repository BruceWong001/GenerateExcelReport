using System;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace GenerateExcelLib
{
    ///
    /// This class is for drawing a excel for multiple data table.
    ///
    public class ExportExcelUtility
    {
        private IExcelWorkBook m_WorkBook;

        public ExportExcelUtility(IExcelWorkBook workBook)
        {
            m_WorkBook=workBook;
        }

        public void GenerateExcel<T>(List<T> Data)
        {
            Boolean isFirstBlock=true; // only the first block need to draw with header.
            foreach(T item in Data)
            {
                using(var dataDesigner=new ExportDataDesigner<T>(item))
                {                    
                    //generate datatable
                    DataTable m_data=dataDesigner.GeneratDataTable();
                    var parameter=new DrawParameter(){
                        StartCol=1,StartRow=1,
                        MergeCells=dataDesigner.MergeCells,
                        HiddenColumns=dataDesigner.HiddenCols
                    };
                    m_WorkBook.DrawExcel(m_data,parameter,isFirstBlock);
                    isFirstBlock=false;                    
                }
            }
            m_WorkBook.Save();
        }

    }
}