using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using Aspose.Cells;

namespace GenerateExcelLib.Tests
{
    public class Test_ExportExcelUtility:IDisposable
    {
        class Learner
        {
            public string Name{get;set;}
            public int Age {get;set;}
        }
        class SessionTime2Elements
        {
            public DateTime Session{get;set;}
            public List<Learner> Learners{get;set;}
            public string Address{get;set;}
        }
        class ListMiddle
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}            
            public List<SessionTime2Elements> Sessions {get;set;}
            public string Trainer {get;set;}
            
        }
        public void Dispose()
        {
            // release resource if you use them during test.
        }

        [Fact]
        [Trait("Category","ExcelUtility")]
        public void Export_Corrent_RowCol()
        {
            //Given
            ListMiddle data1=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="bbb",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}}};
            ListMiddle data2=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="bbb",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}}};
            
            //When
           // using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
            using MemoryStream ms=new MemoryStream();
            var tool=new ExportExcelUtility(new ExportRegularExcel(ms));
            tool.GenerateExcel<ListMiddle>(new List<ListMiddle>{data1,data2});
            
            //Then assert
            var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
            Assert.Equal(7,result.Item1);
            Assert.Equal(11,result.Item2); //10 data rows + 1 header
        }

    }
}