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
        // the format of key is 'start_ColIndex-start_RowIndex' when drill down current Data, since the merge should happen on same level and same value on same column.
        private Dictionary<string,Tuple<int,int,int,int>> m_MergeCells=new Dictionary<string,Tuple<int, int, int, int>>();
        //The data which need to conver to DataTable, it's a generic type, you can define by yourself,
        //if you have array or collection type property in your data definition, please use List<T>. 
        private T Data;
        ///
        /// to record all col index and identifier name as a Key. merge follower could find extension content by identifier column index, then combine the content for comparing logic of merged cell.
        /// one constrain is the identifier must be indicates on previous position than follower.
        ///
        private Dictionary<string,int> m_MergeIdentifiers=new Dictionary<string, int>();
        private const string COlEXTENSION_Name="MergeIdentifier";
        private List<int> m_HiddenCols=new List<int>();

        public Dictionary<string,Tuple<int,int,int,int>> MergeCells{get{
            return m_MergeCells;
        }}
        public List<int> HiddenCols{
            get{return m_HiddenCols;}
        }


        public ExportDataDesigner(T data)
        {
            Data=data;
        }
        public void Dispose()
        {
            m_DT?.Dispose();

        }
        private StructType PreparePropertyInfo(PropertyInfo propertyItem)
        {
            var IsGenericType = propertyItem.PropertyType.IsGenericType;
            var IsBasicType= propertyItem.PropertyType.IsPrimitive || propertyItem.PropertyType.Equals(typeof(String)) || 
                            propertyItem.PropertyType.Equals(typeof(string)) || propertyItem.PropertyType.Equals(typeof(DateTime));

            var list = propertyItem.PropertyType.GetInterface("IEnumerable", false); //retrieve the collection object.

            if (IsGenericType && list != null) return StructType.GenericList;
            if (IsBasicType) 
                return StructType.BasicType;
            else
            {
                return StructType.ComplexType;
            }
            
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
                switch(PreparePropertyInfo(property_Item)){
                    case StructType.BasicType:{
                        if(needAddCol)
                        {
                            var newCol = m_DT.Columns.Add(property_Item.Name, property_Item.PropertyType); //add column in Data table
                            ParseColumnAttribute(currentCol, property_Item, newCol);
                        }

                        var Value=property_Item.GetValue(_data); // add row value
                        row[currentCol]= Value;    //set current cell's value
                        CopyValuetoBelowRows_ForOneCol(currentCol,currentRowNum,Value);//copy current column's value to all below rows.
                        currentCol++;
                         break;
                    }

                    case StructType.GenericList:{
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
                            currentCol= DrillDown(item,start_Col,currentRowNum,isFirstTime_inLoop,isResueRow_Forbelowloop?row:null);
                            isFirstTime_inLoop=false;
                            isResueRow_Forbelowloop=false;
                        }
                        break;                        
                    }
                    default: {
                        // if curren property is class object, it should keep drilling down.
                        currentCol= DrillDown(property_Item.GetValue(_data),currentCol,currentRowNum,needAddCol,row);
                        continue; //since layout all properties from your customized Class, so no need plus column index again.  
                    }
                }
            }
            return currentCol; //return the next column index number.
        }

        private void ParseColumnAttribute(int currentCol, PropertyInfo property_Item, DataColumn newCol)
        {
            var identifierCol = property_Item.GetCustomAttribute<MergeIdentifierAttribute>();
            if (identifierCol != null)
            {//set merge identifier collection.
                m_MergeIdentifiers.Add(identifierCol.Name, currentCol);
                m_HiddenCols.Add(currentCol);
                
            }
            var followerCol = property_Item.GetCustomAttribute<MergeFollowerAttribute>();
            if (followerCol != null)
            {
                newCol.ExtendedProperties.Add(COlEXTENSION_Name, followerCol.IdentifierName);
            }
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
                for(int colIndex=0;colIndex<currentColNumber;colIndex++)//m_DT.Columns.Count
                {
                    // only need to copy the columns value before current column.
                        newRow[colIndex]=oldRow[colIndex];
                        // copy row means all columns in row should be merged from beginning.
                        //the currentColNumber records when trigger the copy event. it means, before this column all copy the cells should be merged.
                        if(m_DT.Columns.Count>1)
                    {//by default, first column should not be needed to merge, so ignore it.
                     //UpdateMergeCoordinate(currentRowNumber,colIndex);
                        Tuple<int, int> ret = FindMergeRows(colIndex, currentRowNumber, newRow[colIndex]); // To find same value rowindex and row count.
                        GenerateMergeCoordinate(colIndex, ret);
                    }
                }                
            }
        }
        void GenerateMergeCoordinate(int colIndex, Tuple<int, int> ret)
        {
            if (ret.Item2 > 1)
            {
                //when merged cell count >1, then it should be merged.
                string key = $"{colIndex}-{ret.Item1}";
                if (m_MergeCells.ContainsKey(key))
                {
                    m_MergeCells[key] = new Tuple<int, int, int, int>(colIndex, ret.Item1, 1, ret.Item2);
                }
                else
                {
                    m_MergeCells.Add(key, new Tuple<int, int, int, int>(colIndex, ret.Item1, 1, ret.Item2));
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
                {// copy all value from current row to the end of table
                    int mergeCount=1;
                    int merge_StartRow=currentRow;
                    //if current column is follower, then find identify column number for the follower column
                    string identifier_Key=m_DT.Columns[colIndex].ExtendedProperties.ContainsKey(COlEXTENSION_Name)?m_DT.Columns[colIndex].ExtendedProperties[COlEXTENSION_Name].ToString():string.Empty;
                    int identtifier_ColNum=m_MergeIdentifiers.ContainsKey(identifier_Key)?m_MergeIdentifiers[identifier_Key]:-1;
                    string Prefix_currentContent=string.Empty;
                    string currentValue=string.Empty;
                    if(identtifier_ColNum>-1)
                    {
                        Prefix_currentContent= m_DT.Rows[currentRow][identtifier_ColNum].ToString();
                    }
                    currentValue=$"{Prefix_currentContent}-{value?.ToString()}";
                    for(int rowNum=currentRow+1;rowNum<m_DT.Rows.Count;rowNum++)
                    {
                        m_DT.Rows[rowNum][colIndex]=value; //set value on all rows but same column.
                        //combine relevant value from required other cell in same row. then compare final content if it's same
                        string Prefix_comparedContent=string.Empty;
                        if(identtifier_ColNum>-1)
                        {
                            Prefix_comparedContent=m_DT.Rows[rowNum][identtifier_ColNum].ToString();
                        }                 
                        string ComparedValue=$"{Prefix_comparedContent}-{value?.ToString()}";

                        if(currentValue.Equals(ComparedValue))//compare is true
                        {                            
                            mergeCount++;
                            if(mergeCount>1)
                            {
                                GenerateMergeCoordinate(colIndex,new Tuple<int, int>(merge_StartRow,mergeCount));
                            }
                        }
                        else
                        {
                            mergeCount=1;
                            merge_StartRow=rowNum;
                            currentValue=ComparedValue;
                        }

                    }
                }
                else if(currentRow==m_DT.Rows.Count-1)
                {//current row is the last row, so it should find above all same value's rows for current column
                    Tuple<int,int> ret= FindMergeRows(colIndex,currentRow,value); // To find same value rowindex and row count.
                    GenerateMergeCoordinate(colIndex, ret);
                }
            }
            
        }
        ///
        /// fine merge cells and return start row number. the condition is for certain column.
        ///
        private Tuple<int, int> FindMergeRows(int colIndex, int currentRow, object value)
        {
            int startRow=currentRow;
            int mergeCount=1;
            //if current column is follower, then find identify column number for the follower column
            string identifier_Key=m_DT.Columns[colIndex].ExtendedProperties.ContainsKey(COlEXTENSION_Name)?m_DT.Columns[colIndex].ExtendedProperties[COlEXTENSION_Name].ToString():string.Empty;
            int identtifier_ColNum=m_MergeIdentifiers.ContainsKey(identifier_Key)?m_MergeIdentifiers[identifier_Key]:-1;
            string Prefix_currentContent=string.Empty;
            string currentValue=string.Empty;
            if(identtifier_ColNum>-1)
            {
                Prefix_currentContent= m_DT.Rows[currentRow][identtifier_ColNum].ToString();
            }
            currentValue=$"{Prefix_currentContent}-{value?.ToString()}";
            for(int rowIndex=currentRow-1;rowIndex>=0;rowIndex--)
            {
                //combine relevant value from required other cell in same row. then compare final content if it's same
                string Prefix_comparedContent=string.Empty;
                if(identtifier_ColNum>-1)
                {
                    Prefix_comparedContent=m_DT.Rows[rowIndex][identtifier_ColNum].ToString();
                }                 
                string ComparedValue=$"{Prefix_comparedContent}-{m_DT.Rows[rowIndex][colIndex].ToString()}";
                //
                if(currentValue.Equals(ComparedValue))
                {
                    startRow=rowIndex;
                    mergeCount++;
                }
                else
                    break;
                
            }
            
            return new Tuple<int, int>(startRow,mergeCount);
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