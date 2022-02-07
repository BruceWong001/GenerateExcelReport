using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;
using Aspose.Cells;


namespace GenerateExcelLib.Tests
{
    public class Learner
    {
        public string Name{get;set;}
        public int Age {get;set;}
    }
    public class SessionTime
    {
        public DateTime Session{get;set;}
    }
    public class MyClass
    {
        public string ClassTitle{get;set;}
        public string ClassCode{get;set;}
        public List<DateTime> Sessions {get;set;}
        public string Trainer {get;set;}
        public List<Learner> learners {get;set;}
    }
    public class SimpleClass
    {
        public string ClassTitle{get;set;}
        public string ClassCode{get;set;}
        public string Trainer {get;set;}
    }
    public class Test_ExportDataDesigner
    {
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_SimpleObj()
        {
            // Given
            SimpleClass data=new SimpleClass(){ClassTitle="Java",ClassCode="10010",Trainer="Bill"};
            // When
            var designer=new ExportDataDesigner<SimpleClass>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(1,table.Rows.Count);
            Assert.Equal(3,table.Columns.Count);
            Assert.Equal("Bill",table.Rows[0][2].ToString());
            
        }
 
        public class ListStart
        {
            public List<SessionTime> Sessions {get;set;}
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}
        }        
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtBegin()
        {
            // Given
            ListStart data=new ListStart(){Sessions=new List<SessionTime>(){new SessionTime{Session=DateTime.Now},new SessionTime{Session=DateTime.Now.AddDays(1)}},
                                     ClassTitle="Java",ClassCode="10010"};
            // When
            var designer=new ExportDataDesigner<ListStart>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(2,table.Rows.Count);
            Assert.Equal(3,table.Columns.Count);
            Assert.Equal("Java",table.Rows[0][1].ToString());
        }
        public class ListEnd
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}
            public string Trainer {get;set;}
            public List<Learner> learners {get;set;}
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtEnd()
        {
            // Given
            ListEnd data=new ListEnd(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       learners= new List<Learner>{new Learner{Name="Lily",Age=20},new Learner{Name="Joe",Age=19},new Learner{Name="Wuli",Age=28}}};
            // When
            var designer=new ExportDataDesigner<ListEnd>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(3,table.Rows.Count);
            Assert.Equal(5,table.Columns.Count);
            Assert.Equal("10010",table.Rows[0][1].ToString());
        }

    }
}