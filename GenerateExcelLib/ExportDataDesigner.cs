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
        private int DrillDown(Object _data,int indexCol,Boolean needAddCol=true,DataRow parentRow=null)
        {
            DataRow row;
            if(parentRow is null)
            {
                row=m_DT.NewRow();
                m_DT.Rows.Add(row);
            }
            else{
                row=parentRow;
            }
            var newType = _data.GetType();
            foreach (var property_Item in newType.GetRuntimeProperties())
            {                
                var propertyName = property_Item.Name;
                var IsGenericType = property_Item.PropertyType.IsGenericType;
                var list = property_Item.PropertyType.GetInterface("IEnumerable", false);
                if (IsGenericType && list != null)
                {
                    // if current property is  a list type.
                    var listVal = property_Item.GetValue(_data) as IEnumerable<object>;
                    if (listVal == null) 
                    {
                        continue;
                    }
                    int start_Col=indexCol; // record the column index for the looping of List items
                    Boolean isFirst=needAddCol;
                    foreach (var item in listVal)
                    {         
                        DrillDown(item,start_Col,isFirst,isFirst?row:null);
                        isFirst=false;
                    }
                }
                else
                {
                    if(needAddCol)
                    {
                        m_DT.Columns.Add(propertyName,property_Item.PropertyType); //add column in Data table
                    }
                    
                    var Value=property_Item.GetValue(_data); // add row value
                    row[indexCol]= Value;
                }
                indexCol++;
            }
            return indexCol; //return the next column index number.
        }
        public DataTable GeneratDataTable()
        {
            //
            DrillDown(Data,0);
            return m_DT;

        }

    }

}