using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace GenerateExcelLib.Tests
{

    public class Test_ClosedXML:IDisposable
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
        [Trait("Category","ClosedXML")]
        public void Export_OneDataTable_withHead()
        {
            //Arrange: generate datatable
            using(DataTable mydata=Initial_Simple_DataTable())
            {
                //using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcelClosedXML(ms);
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
        [Trait("Category","ClosedXML")]
        public void Export_OneDataTable_withNoHead()
        {
            using(DataTable mydata=Initial_Simple_DataTable())
            {
               // using(FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create))
                using(MemoryStream ms=new MemoryStream())
                {
                    var work_book=new ExportRegularExcelClosedXML(ms);
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





    }

}