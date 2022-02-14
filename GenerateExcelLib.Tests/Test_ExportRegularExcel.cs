using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;
using Aspose.Cells;

namespace GenerateExcelLib.Tests
{

    public class Test_ExportExcel:IDisposable
    {
        private DataTable Initial_Simple_DataTable()
        {
            DataTable mydata=new DataTable("Table1");

            mydata.Columns.Add("Class Code",typeof(string));
            mydata.Columns.Add("Class Title",typeof(string));
            mydata.Columns.Add("Time Slot",typeof(DateTime));

            mydata.Rows.Add("C-01-1001","PA Class 1",DateTime.Now);
            mydata.Rows.Add("C-01-1002","PA Class 2",DateTime.Now.AddHours(1));
            mydata.Rows.Add("C-01-1003","PA Class 3",DateTime.Now.AddHours(2));
            return mydata;

        }
        public void Dispose()
        {
            // release resource if you use them during test.
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_CorrectColumns_withHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();
                //Act: run test function
                using( Workbook result_workbook= work_book.GenerateExcel(mydata))
                {
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream). test only
                        result_workbook.Save(ms,SaveFormat.Xlsx);

                        //Assert: result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(3,result.Item1);
                        Assert.Equal(4,result.Item2);
                    }
                }
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_CorrectColumns_withNoHead()
        {
            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();

                //run test function
                using( Workbook result_workbook= work_book.GenerateExcel(mydata,1,false))
                {
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        result_workbook.Save(ms,SaveFormat.Xlsx);
                    //  result_workbook.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(3,result.Item1); //assert column num
                        Assert.Equal(3,result.Item2); //assert row num
                    }
                }
            }
            
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_CorrectColumns_withNoHead_start_RowNum()
        {
            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();

                //run test function
                using( Workbook result_workbook= work_book.GenerateExcel(mydata,5,false))
                {
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        result_workbook.Save(ms,SaveFormat.Xlsx);
                        //result_workbook.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(3,result.Item1); //assert column num
                        Assert.Equal(7,result.Item2); //assert row num
                    }
                }
            }
            
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_CorrectContent_withHead()
        {
            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();
                //run test function
                using( Workbook result_workbook= work_book.GenerateExcel(mydata))
                {
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        result_workbook.Save(ms,SaveFormat.Xlsx);
                        //result_workbook.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,4);
                        Assert.Equal("C-01-1003",result); //assert certain cell's value.

                    }
                }
            }
            
        }    
        [Fact]
        [Trait("Category","Basic")]
        public void Export_CorrectContent_withNoHead()
        {

            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();
                //run test function
                using( Workbook result_workbook= work_book.GenerateExcel(mydata,1,false))
                {
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        result_workbook.Save(ms,SaveFormat.Xlsx);
                        //result_workbook.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,2);
                        Assert.Equal("PA Class 2",result); //assert certain cell's value.

                    }
                }
            }
            
        }           
        [Fact]
        [Trait("Category","Basic")]
        public void Export_MergeCell_Multi_Rows()
        {
            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();
                using(var merge_Book= work_book.GenerateExcel(mydata))
                {
                    //run test function
                    work_book.MergeCell(merge_Book,2,2,1,2);
                
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        merge_Book.Save(ms,SaveFormat.Xlsx);
                       // merge_Book.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Is_MergeCell(ms,2,2,1,2);
                        Assert.True(result); //assert the spicified area is merged.

                    }
                }
            }
            
        }   

        [Fact]
        [Trait("Category","Basic")]
        public void Export_MergeCell_Multi_Cols()
        {
            //generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                var work_book=new ExportRegularExcel();
                using(var merge_Book= work_book.GenerateExcel(mydata))
                {
                    //run test function
                    work_book.MergeCell(merge_Book,1,3,2,1);
                
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        merge_Book.Save(ms,SaveFormat.Xlsx);
                      //  merge_Book.Save(@"c:\test.xlsx"); // only for debug
                        //Assert result
                        var result=Excel_Ops_Aspose.Is_MergeCell(ms,1,3,2,1);
                        Assert.True(result); //assert the spicified area is merged.

                    }
                }
            }
            
        }    
        [Fact]
        [Trait("Category","Basic")]
        public void Export_MergeCell_withArgumentException()
        {
            var work_book=new ExportRegularExcel();
            
            //run test function
            var exception= Assert.Throws<ArgumentException>(()=> work_book.MergeCell(null,0,0,1,2));


        }      
        [Fact]
        [Trait("Category","Basic")]
        public void Export_MergeCell_withNullObjectException()
        {
            var work_book=new ExportRegularExcel();
            

            //run test function
            var exception= Assert.Throws<NullReferenceException>(()=> work_book.MergeCell(null,1,1,1,2));


        }
        [Fact(Skip="demo skip")]
        public void Export_MergeCell_demoSkip()
        {
            var work_book=new ExportRegularExcel();
            

            //run test function
            var exception= Assert.Throws<NullReferenceException>(()=> work_book.MergeCell(null,1,1,1,2));


        }
                 
    }

}