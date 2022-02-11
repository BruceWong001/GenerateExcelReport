using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections.Generic;


namespace GenerateExcelLib
{
    ///
    /// export data structure, defined by end user.
    /// Note: all export field need to be defined via property and public as well.
    /// Note: the index of column and row should be start from 0, not 1.
    ///    
    public class ExportDataDesigner<T>:IDisposable
    {
        // final designed data table. all columns and rows will reuse this object. 
        private DataTable m_DT=new DataTable(); 
        // item's order in Tuple are first Col, first Row,total Cols, total Rows
        // the format of key is 'recursion level-ColIndex' when drill down current Data, since the merge should happen on same level and same value on same column.
        private Dictionary<string,Tuple<int,int,int,int>> m_MergeCells=new Dictionary<string,Tuple<int, int, int, int>>();
        //The data which need to conver to DataTable, it's a generic type, you can define by yourself,
        //if you have array or collection type property in your data definition, please use List<T>. 
        private T Data;
        public Dictionary<string,Tuple<int,int,int,int>> MergeCells{get{
            return m_MergeCells;
        }}

        public ExportDataDesigner(T data)
        {
            Data=data;
        }
        public void Dispose()
        {
            m_DT?.Dispose();

        }
        ///
        /// recursion method. list all sub collection into one row.
        ///
        private int DrillDown(Object _data,int currentCol,int currentRowIndex=0,Boolean needAddCol=true,DataRow reuse_Row=null)
        {
            DataRow row;
            int currentRowNum=currentRowIndex; //record current row number for copy data method.
            if(reuse_Row is null)
            { //create a new row in datatable.
                row=m_DT.NewRow();
                DataRow previousRow=m_DT.Rows.Count>0?m_DT.Rows[m_DT.Rows.Count-1]:null;
                m_DT.Rows.Add(row); 
                currentRowNum=m_DT.Rows.Count-1;
                CopyRowValuefromAboveRow(currentCol,currentRowNum,row,previousRow); 
            }
            else{
                row=reuse_Row;
            }
            var newType = _data.GetType();
            foreach (var property_Item in newType.GetRuntimeProperties())
            {                
                var propertyName = property_Item.Name;
                var IsGenericType = property_Item.PropertyType.IsGenericType;
                var IsBasicType= property_Item.PropertyType.IsPrimitive || property_Item.PropertyType.Equals(typeof(String)) || 
                                property_Item.PropertyType.Equals(typeof(string)) || property_Item.PropertyType.Equals(typeof(DateTime));
                var list = property_Item.PropertyType.GetInterface("IEnumerable", false); //retrieve the collection object.
                if (IsGenericType && list != null)
                { // if current property is  a list type.
                    var listVal = property_Item.GetValue(_data) as IEnumerable<object>;
                    if (listVal == null) continue;
                    int start_Col=currentCol; // record the column index for the looping of List items
                    Boolean isFirstTime_inLoop=needAddCol; //indicate if below loop needs to add new columns.
                    Boolean isResueRow_Forbelowloop=true; //indicate below loop if it needs to add a new row or reuse existing one.
                    foreach (var item in listVal)
                    {   /*
                         (isFirstTime_inLoop?row:null) this condition means it's firs time to drill down the first item in one list,
                         but you must reuse DataRow object, since you have already create this row on top.
                        */ 
                        DrillDown(item,start_Col,currentRowNum,isFirstTime_inLoop,isResueRow_Forbelowloop?row:null);
                        isFirstTime_inLoop=false;
                        isResueRow_Forbelowloop=false;
                    }
                }
                else if(!IsBasicType)
                {// if curren property is data object, it should keep drilling down.
                    currentCol= DrillDown(property_Item.GetValue(_data),currentCol,currentRowNum,true,row);
                    continue; //since layout all properties from your customized Class, so no need plus column index again.  
                }
                else
                {
                    if(needAddCol)
                    {
                        m_DT.Columns.Add(propertyName,property_Item.PropertyType); //add column in Data table
                    }
                    
                    var Value=property_Item.GetValue(_data); // add row value
                    row[currentCol]= Value;
                    CopyValuetoBelowRows_ForOneCol(currentCol,currentRowNum,Value);//copy current column's value to all below rows.
                }
                currentCol++;
            }
            return currentCol; //return the next column index number.
        }
        ///
        /// only copy the latest row's data to curren row.
        ///
        private void CopyRowValuefromAboveRow(int currentColNumber,int currentRowNumber,DataRow newRow,DataRow oldRow)
        {
            int rowCount=m_DT.Rows.Count;
            if(rowCount>1 && oldRow is not null)
            {
                //only handle there is existing data.
                for(int colIndex=0;colIndex<m_DT.Columns.Count;colIndex++)
                {
                    newRow[colIndex]=oldRow[colIndex];
                    // copy row means all columns in row should be merged from beginning.
                    if(colIndex< currentColNumber)
                    {
                        //the currentColNumber records when trigger the copy event. it means, before this column all copy the cells should be merged.
                        if(m_DT.Columns.Count>currentColNumber+1)
                        {//by default, first column should not be needed to merge, so ignore it.
                            UpdateMergeCoordinate(currentRowNumber,colIndex);
                        }
                    }
                }
                
            }
        }
        ///
        /// only copy current cell value to all below rows with same column
        ///
        private void CopyValuetoBelowRows_ForOneCol(int colIndex,int rowIndex,object value)
        {
            int currentRow=rowIndex;

            if(m_DT.Rows.Count>1)
            {
                if(currentRow<m_DT.Rows.Count-1)
                {// copy all value for entire column
                    for(int rowNum=currentRow+1;rowNum<m_DT.Rows.Count;rowNum++)
                    {
                        m_DT.Rows[rowNum][colIndex]=value; //set value on all rows but same column.
                    }
                    //current cloumn from row 0 should be merged.
                    string key=$"{colIndex}-{0}";               
                    if(m_MergeCells.ContainsKey(key))
                    {
                        Tuple<int,int,int,int> originOne=m_MergeCells[key];
                        m_MergeCells[key]=new Tuple<int, int, int, int>(colIndex,0,1,m_DT.Rows.Count);
                    }
                    else
                    {
                        m_MergeCells.Add(key,new Tuple<int, int, int, int>(colIndex,0,1,m_DT.Rows.Count));
                    }
                }
                else if(currentRow==m_DT.Rows.Count-1)
                {//current row is the last row, so it should find above all same value's rows for current column
                    Tuple<int,int> ret= FindMergeRows(colIndex,currentRow,value); // To find same value rowindex and row count.
                    if(ret.Item2>1)
                    {
                        //when merged cell count >1, then it should be merged.
                        //current cloumn from return row should be merged.
                        string key=$"{colIndex}-{ret.Item1}";               
                        if(m_MergeCells.ContainsKey(key))
                        {
                            Tuple<int,int,int,int> originOne=m_MergeCells[key];
                            m_MergeCells[key]=new Tuple<int, int, int, int>(colIndex,ret.Item1,1,ret.Item2);
                        }
                        else
                        {
                            m_MergeCells.Add(key,new Tuple<int, int, int, int>(colIndex,ret.Item1,1,ret.Item2));
                        }
                    }
                }
            }
            
        }

        private Tuple<int, int> FindMergeRows(int colIndex, int currentRow, object value)
        {
            int startRow=currentRow;
            int mergeCount=1;
            for(int rowIndex=currentRow-1;rowIndex>=0;rowIndex--)
            {
                if(m_DT.Rows[rowIndex][colIndex].ToString().Equals(value.ToString()))
                {
                    startRow=rowIndex;
                    mergeCount++;
                }
                else
                    break;
                
            }
            
            return new Tuple<int, int>(startRow,mergeCount);
        }

        private void  UpdateMergeCoordinate(int currentRow,int colIndex)
        { 
            if(currentRow>0)
            {
                string key=$"{colIndex}-{0}";
                if(m_MergeCells.ContainsKey(key))
                {
                    Tuple<int,int,int,int> originOne=m_MergeCells[key];
                    m_MergeCells[key]=new Tuple<int, int, int, int>(colIndex,originOne.Item2,1,originOne.Item4+1);
                }
                else
                {
                    m_MergeCells.Add(key,new Tuple<int, int, int, int>(colIndex,0,1,2));
                }
            }

        }
        ///
        /// the DataTable which return to the caller will be disposed when current ExportDataDesigner dispose.
        /// so no need to dispose explicitly
        ///
        public DataTable GeneratDataTable()
        {
            //
            DrillDown(Data,0,0);
            return m_DT;

        }

    }

}