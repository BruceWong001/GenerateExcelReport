using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

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

 //////////////////////////////       

        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_withHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata);
                    work_book.Save();
                    //Assert: result
                   var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(3,result.Item1);
                   Assert.Equal(4,result.Item2);

                } 
            }
        }
        
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_withNoHead()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(3,result.Item1);
                    Assert.Equal(3,result.Item2);

                }
            }             
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_FromSpecifiedPosition_2()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,1,1);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(3,result.Item1);
                    Assert.Equal(4,result.Item2);

                }
            }             
        }
        [Fact]
        [Trait("Category","Basic")]
       public void Export_OneDataTable_4SpecifiedPosition_2ByParameter()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    DrawParameter p=new DrawParameter{
                        StartCol=1,StartRow=1
                    };
                    //Act: run test function
                    work_book.DrawExcel(mydata,p);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(3,result.Item1);
                    Assert.Equal(4,result.Item2);

                }
            }             
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_FromSpecifiedPosition()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,5,2);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(4,result.Item1);
                    Assert.Equal(8,result.Item2);

                }
            }             
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_4SpecifiedPosition_byparameter()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    DrawParameter p=new DrawParameter{
                        StartRow=5,StartCol=2
                    };
                    //Act: run test function
                    work_book.DrawExcel(mydata,p);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(4,result.Item1);
                    Assert.Equal(8,result.Item2);

                }
            }             
        }
       [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_withNoHead_FromSpecifiedPosition()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,5,2,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(4,result.Item1);
                    Assert.Equal(7,result.Item2);

                }
            }             
        }
       [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_withNoHead_4SpecifiedPosition_byparameter()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    DrawParameter p=new DrawParameter{
                        StartRow=5,StartCol=2
                    };
                    //Act: run test function
                    work_book.DrawExcel(mydata,p,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(4,result.Item1);
                    Assert.Equal(7,result.Item2);

                }
            }             
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_withHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,true);
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                   var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(3,result.Item1);
                   Assert.Equal(7,result.Item2);

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_withHead_byparameter()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    DrawParameter p=new DrawParameter{
                        StartCol=1,StartRow=1
                    };
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,p,true);
                    work_book.DrawExcel(mydata,p,false);
                    work_book.Save();
                    //Assert: result
                   var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(3,result.Item1);
                   Assert.Equal(7,result.Item2);

                } 
            }
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_FromSpeicifiedPosition()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,5,2,true);
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                   var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(4,result.Item1);
                   Assert.Equal(11,result.Item2);

                } 
            }
        }

        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_4SpeicifiedPosition_byparameter()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    DrawParameter p=new DrawParameter{
                        StartRow=5,StartCol=2
                    };
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,p,true);
                    work_book.DrawExcel(mydata,p,false);
                    work_book.Save();
                    //Assert: result
                   var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(4,result.Item1);
                   Assert.Equal(11,result.Item2);

                } 
            }
        }


        [Fact]
        [Trait("Category","Basic")]
        public void Export_ConstructExportObj_withNullStreamException()
        {
            //run test function
            var exception= Assert.Throws<InvalidDataException>(()=> new ExportRegularExcel(null));


        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_Content()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,4);
                    Assert.Equal("C-01-1003",result); //assert certain cell's value.
                    var result1=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,2);
                    Assert.Equal("PA Class 1",result1); //assert certain cell's value.
                } 
            }
        }
       [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_Content_withNoHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,3);
                    Assert.Equal("C-01-1003",result); //assert certain cell's value.
                    var result1=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,1);
                    Assert.Equal("PA Class 1",result1); //assert certain cell's value.

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTable_Content()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata);
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,7);
                    Assert.Equal("C-01-1003",result); //assert certain cell's value.
                    var result1=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,5);
                    Assert.Equal("PA Class 1",result1); //assert certain cell's value.

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTable_Content_withNoHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    //Act: run test function
                    work_book.DrawExcel(mydata,false);
                    work_book.DrawExcel(mydata,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,6);
                    Assert.Equal("C-01-1003",result); //assert certain cell's value.
                    var result1=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,4);
                    Assert.Equal("PA Class 1",result1); //assert certain cell's value.

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_MergeCell()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter,true);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Is_MergeCell(ms,2,3,1,2);
                    Assert.True(result); //assert the spicified area is merged.

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTable_MergeCell_withNoHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter,false);
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Is_MergeCell(ms,2,2,1,2);
                    Assert.True(result); //assert the spicified area is merged.

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_MergeCell()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    #region DataTable 1
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter,true);
                    #endregion
                    #region  DataTable 2
                    var parameter2=new DrawParameter{
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter2,false);
                    #endregion

                    work_book.Save();
                    //Assert: result
                    Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,3,1,2));
                    Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,6,1,2));

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_TwoDataTables_MergeCell_withNoHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    #region DataTable 1
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter,false);
                    #endregion
                    #region  DataTable 2
                    var parameter2=new DrawParameter{
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        }
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter2,false);
                    #endregion

                    work_book.Save();
                    //Assert: result
                    Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,2,1,2));
                    Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,5,1,2));

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTables_RemoveColumn()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
              //  using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        },
                        HiddenColumns=new List<int>{2}
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter);
                
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                   Assert.Equal(2,result.Item1);

                } 
            }
        }
        [Fact]
        [Trait("Category","Basic")]
        public void Export_OneDataTables_RemoveMultiColumn()
        {
            //Arrange: generate datatable
            DataTable mydata=new DataTable("Table1");

            mydata.Columns.Add("Class Code",typeof(string));
            mydata.Columns.Add("Class Title",typeof(string));
            mydata.Columns.Add("Time Slot",typeof(DateTime));
            mydata.Columns.Add("Trainer",typeof(string));

            mydata.Rows.Add("C-01-1001","PA Class 1",DateTime.Now,"Joe");
            mydata.Rows.Add("C-01-1002","PA Class 2",DateTime.Now.AddHours(1),"Lily");
            mydata.Rows.Add("C-01-1003","PA Class 2",DateTime.Now.AddHours(2),"Lucy");

            using(mydata)
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcel(ms);
                    
                    var parameter=new DrawParameter{
                        StartCol=1,StartRow=1,
                        MergeCells=new Dictionary<string, Tuple<int, int, int, int>>{
                            {"1-0",new Tuple<int, int, int, int>(1,1,1,2)}
                        },
                        HiddenColumns=new List<int>{0,2}
                    } ;
                    //Act: run test function
                    work_book.DrawExcel(mydata,parameter);
                
                    work_book.Save();
                    //Assert: result
                    var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                    Assert.Equal(2,result.Item1);
                    var cellContent1=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,1,3);
                    Assert.Equal("PA Class 2",cellContent1); //assert certain cell's value.
                    var cellContent2=Excel_Ops_Aspose.Retrieve_Content_CertainCell(ms,2,3);
                    Assert.Equal("Lily",cellContent2); //assert certain cell's value.

                } 
            }
        }
        
    }

}