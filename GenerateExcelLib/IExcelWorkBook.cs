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
        public Dictionary<string,Tuple<int,int,int,int>> MergeCells{get;set;}
        public List<int> HiddenColumns{get;set;}
        
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