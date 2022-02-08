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
        private int DrillDown(Object _data,int currentCol,Boolean needAddCol=true,DataRow reuse_Row=null)
        {
            DataRow row;
            if(reuse_Row is null)
            {
                //create a new row in datatable.
                row=m_DT.NewRow();
                m_DT.Rows.Add(row); 
            }
            else{
                row=reuse_Row;
            }
            var newType = _data.GetType();
            foreach (var property_Item in newType.GetRuntimeProperties())
            {                
                var propertyName = property_Item.Name;
                var IsGenericType = property_Item.PropertyType.IsGenericType;
                var list = property_Item.PropertyType.GetInterface("IEnumerable", false); //retrieve the collection object.
                if (IsGenericType && list != null)
                {
                    // if current property is  a list type.
                    var listVal = property_Item.GetValue(_data) as IEnumerable<object>;
                    if (listVal == null) 
                    {
                        continue;
                    }
                    int start_Col=currentCol; // record the column index for the looping of List items
                    Boolean isFirstTime_inLoop=needAddCol;
                    foreach (var item in listVal)
                    {
                        /*
                         (isFirstTime_inLoop?row:null) this condition means it's firs time to drill down the first item in one list,
                         but you must reuse DataRow object, since you have already create this row on top.
                        */ 
                        DrillDown(item,start_Col,isFirstTime_inLoop,isFirstTime_inLoop?row:null);
                        isFirstTime_inLoop=false;
                    }
                }
                else
                {
                    if(needAddCol)
                    {
                        m_DT.Columns.Add(propertyName,property_Item.PropertyType); //add column in Data table
                    }
                    
                    var Value=property_Item.GetValue(_data); // add row value
                    row[currentCol]= Value;
                }
                currentCol++;
            }
            return currentCol; //return the next column index number.
        }
        public DataTable GeneratDataTable()
        {
            //
            DrillDown(Data,0);
            return m_DT;

        }

    }

}