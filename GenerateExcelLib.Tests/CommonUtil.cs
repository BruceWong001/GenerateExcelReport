
using System;
using System.IO;
using Aspose.Cells;
using System.Collections.Generic;
using System.Collections;

namespace GenerateExcelLib.Tests
{
    public class Excel_Ops_Aspose
    {
        public static Tuple<int,int> Retrieve_Num_Column_Row(Stream _stream)
        {
            if(_stream!=null)
            {
                _stream.Position=0; //return the point of stream back to the beginning.
                
                using(Workbook workbook = new Workbook(_stream))
                {
                    Cells cells = workbook.Worksheets[0].Cells;
                    //Note: if the exact column number is 3 but the cells.MaxColumn=2. the MaxRow is same behavior.
                    return new Tuple<int, int>(cells.MaxColumn+1,cells.MaxRow+1); 
                }
            
            }
            else
            {
                return null;
            }
        
            
        }
        public static string Retrieve_Content_CertainCell(Stream _stream, int Col,int Row )
        {
            if(_stream!=null)
            {
                _stream.Position=0; //return the point of stream back to the beginning.
                
                using(Workbook workbook = new Workbook(_stream))
                {
                    Cells cells = workbook.Worksheets[0].Cells;
                    //Note: if the exact column number is 3 but the cells.MaxColumn=2. the MaxRow is same behavior.
                    return cells[Row-1,Col-1].Value.ToString();
                }
            
            }
            else
            {
                return string.Empty;
            }

        }
        public static Boolean Is_MergeCell(Stream _stream, int startCol,int startRow,int totalCols,int totalRows )
        {
            if(_stream!=null)
            {
                _stream.Position=0; //return the point of stream back to the beginning.
                
                using(Workbook workbook = new Workbook(_stream))
                {
                    Cells cells = workbook.Worksheets[0].Cells;
                    // check if the region which is between col and row is merged.
                    ArrayList mergedlist= cells.MergedCells;
                    if(mergedlist.Count>0)
                    {
                        int col_range=totalCols-1; // merge column range
                        int row_range=totalRows-1; // merge row range
                        int exact_startcol=startCol-1;
                        int exact_startrow=startRow-1;
                        foreach(CellArea mergedCell in mergedlist)
                        {
                            if(mergedCell.StartColumn==(startCol-1) && mergedCell.StartRow==(startRow-1) &&
                                    mergedCell.EndColumn==exact_startcol+col_range && mergedCell.EndRow==exact_startrow+row_range)
                                    {
                                        return true;
                                    }
                        }
                    }
                    
                    return false;
                    
                }
            
            }
            else
            {
                return false;
            }

        }
    }
}