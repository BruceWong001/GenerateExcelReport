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
    ///    
    public class ExportDataDesigner<T>:IDisposable
    {
        // final designed data table. all columns and rows will reuse this object. 
        private DataTable m_DT=new DataTable(); 
        // item's order in Tuple are first Col, first Row,total Cols, total Rows
        private List<Tuple<int,int,int,int>> m_MergeCells=new List<Tuple<int, int, int, int>>();
        //The data which need to conver to DataTable, it's a generic type, you can define by yourself,
        //if you have array or collection type property in your data definition, please use List<T>. 
        public T Data {get;set;}

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
            {
                //create a new row in datatable.
                row=m_DT.NewRow();
                DataRow previousRow=m_DT.Rows.Count>0?m_DT.Rows[m_DT.Rows.Count-1]:null;
                m_DT.Rows.Add(row); 
                currentRowNum=m_DT.Rows.Count-1;
                CopyRowValuefromAboveRow(row,previousRow); 
            }
            else{
                row=reuse_Row;
            }
            var newType = _data.GetType();
            foreach (var property_Item in newType.GetRuntimeProperties())
            {                
                var propertyName = property_Item.Name;
                var IsGenericType = property_Item.PropertyType.IsGenericType;
                var IsBasicType= property_Item.PropertyType.IsPrimitive || property_Item.PropertyType.Equals(typeof(String)) || property_Item.PropertyType.Equals(typeof(string)) || 
                                    property_Item.PropertyType.Equals(typeof(DateTime));
                var list = property_Item.PropertyType.GetInterface("IEnumerable", false); //retrieve the collection object.
                if (IsGenericType && list != null)
                {
                    // if current property is  a list type.
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
                    CopyValuetoBelowRows(currentCol,currentRowNum,Value);//copy new column's value to all above rows.
                }
                currentCol++;
            }
            return currentCol; //return the next column index number.
        }
        ///
        /// only copy the latest row's data to curren row.
        ///
        private void CopyRowValuefromAboveRow(DataRow newRow,DataRow oldRow)
        {
            int rowCount=m_DT.Rows.Count;
            if(rowCount>1 && oldRow is not null)
            {
                //only handle there is existing data.
                for(int colIndex=0;colIndex<m_DT.Columns.Count;colIndex++)
                {
                    newRow[colIndex]=oldRow[colIndex];
                }
                
            }
        }
        ///
        ///only copy current cell value to all above rows
        ///
        private void CopyValuetoBelowRows(int colIndex,int rowIndex,object value)
        {
            int currentRow=rowIndex;
            if(m_DT.Rows.Count>1 && currentRow<m_DT.Rows.Count-1)
            {
                for(int rowNum=currentRow+1;rowNum<m_DT.Rows.Count;rowNum++)
                {
                    m_DT.Rows[rowNum][colIndex]=value;
                }
            }
            
        }
        public DataTable GeneratDataTable()
        {
            //
            DrillDown(Data,0,0);
            return m_DT;

        }

    }

}