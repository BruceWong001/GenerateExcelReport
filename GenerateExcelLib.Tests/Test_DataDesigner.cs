using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;
using Aspose.Cells;


namespace GenerateExcelLib.Tests
{
    class Learner
    {
        public string Name{get;set;}
        public int Age {get;set;}
    }
    class SessionTime
    {
        public DateTime Session{get;set;}
    }

    public class Test_ExportDataDesigner
    {
        class SimpleClass
        {
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}
            public string Trainer {get;set;}
        }
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
        class SimpleClassEx
        {
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}
            public Learner Student{get;set;}
            public string Trainer {get;set;}
            public DateTime RegistryTime{get;set;}
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_SimpleObj_WithObjMember()
        {
            // Given
            SimpleClassEx data=new SimpleClassEx(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",Student=new Learner{Name="Bruce",Age=30},RegistryTime=DateTime.Now};
            // When
            var designer=new ExportDataDesigner<SimpleClassEx>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(1,table.Rows.Count);
            Assert.Equal(6,table.Columns.Count);
            Assert.Equal("Bill",table.Rows[0][4].ToString());
            
        }         
        class ListStart
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
        }
        class ListEnd
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

        }
        class ListMiddle
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}            
            public List<SessionTime> Sessions {get;set;}
            public string Trainer {get;set;}
            
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtMiddle()
        {
            // Given
            ListMiddle data=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime>{new SessionTime{Session=DateTime.Now},new SessionTime{Session=DateTime.Now.AddDays(1)},new SessionTime{Session=DateTime.Now.AddDays(2)}}};
            // When
            var designer=new ExportDataDesigner<ListMiddle>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(3,table.Rows.Count);
            Assert.Equal(4,table.Columns.Count);

        }
        class SessionObj
        {
            public DateTime Session {get;set;}
            public List<Learner> Learners {get;set;}
        }
        class ComprehensiveObj
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}            
            public string Trainer {get;set;}
            public List<SessionObj> SessionList{get;set;}
            
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_Validate_ColRow_Num()
        {
            // Given
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Bruce",Age=30}}}}                            
                        
                       };
            // When
            var designer=new ExportDataDesigner<ComprehensiveObj>(data);
            DataTable table=designer.GeneratDataTable();
            // Then
            Assert.Equal(3,table.Rows.Count);
            Assert.Equal(6,table.Columns.Count);

        }

    }
}