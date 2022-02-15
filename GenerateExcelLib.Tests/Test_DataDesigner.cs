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
    class SessionTime
    {
        public DateTime Session{get;set;}
        public string SessionName{get;set;}
    }
    class SessionTime2Elements
    {
        public DateTime Session{get;set;}
        public List<Learner> Learners{get;set;}
        public string Address{get;set;}
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
            using(var designer=new ExportDataDesigner<SimpleClass>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(1,table.Rows.Count);
                Assert.Equal(3,table.Columns.Count);
                Assert.Equal("Bill",table.Rows[0][2].ToString());
            }
            
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
            using(var designer=new ExportDataDesigner<SimpleClassEx>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(1,table.Rows.Count);
                Assert.Equal(6,table.Columns.Count);
                Assert.Equal("Bill",table.Rows[0][4].ToString());
            }
            
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
            ListStart data=new ListStart(){Sessions=new List<SessionTime>(){new SessionTime{Session=DateTime.Now,SessionName="class1"},new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="class2"}},
                                     ClassTitle="Java",ClassCode="10010"};
            // When
            using(var designer=new ExportDataDesigner<ListStart>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(2,table.Rows.Count);
                Assert.Equal(4,table.Columns.Count);
            }
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtBegin()
        {
            // Given
            ListStart data=new ListStart(){Sessions=new List<SessionTime>(){new SessionTime{Session=DateTime.Now,SessionName="class1"},new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="class2"}},
                                     ClassTitle="Java",ClassCode="10010"};
            // When
            using(var designer=new ExportDataDesigner<ListStart>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal("10010",table.Rows[0][3].ToString());
                Assert.Equal("10010",table.Rows[1][3].ToString());
                Assert.Equal("Java",table.Rows[0][2].ToString());
                Assert.Equal("Java",table.Rows[1][2].ToString());
            }
        }        
        [Fact]
        [Trait("Category","MergedCells")]
        public void ValidateMergedCells_AtBegin()
        {
            // Given
            ListStart data=new ListStart(){Sessions=new List<SessionTime>(){new SessionTime{Session=DateTime.Now,SessionName="aaa"},new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="bbb"}},
                                     ClassTitle="Java",ClassCode="10010"};
            // When
            using(var designer=new ExportDataDesigner<ListStart>(data))
            {
                DataTable table=designer.GeneratDataTable();
                var mergeCells=designer.MergeCells;
                // Then
                Assert.Equal(2,mergeCells.Count);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(2,0,1,2),mergeCells["2-0"]);

                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(3,0,1,2),mergeCells["3-0"]);
            }
        }    
        class ListEnd
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}
            public string Trainer {get;set;}
            public List<SessionTime2Elements> Sessions {get;set;}
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtEnd()
        {
            // Given
            ListEnd data=new ListEnd(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now,Address="aaa", Learners=new List<Learner>{ new Learner{Name="Lily",Age=20}}},
                       new SessionTime2Elements{ Session=DateTime.Now.AddDays(1),Address="bbbb", Learners=new List<Learner>{ new Learner{Name="Joe",Age=19},new Learner{Name="Bruce",Age=20}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc", Learners=new List<Learner>{ new Learner{Name="Wuli",Age=28}}}}};
            // When
            using(var designer=new ExportDataDesigner<ListEnd>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(4,table.Rows.Count);
                Assert.Equal(7,table.Columns.Count);
            }

        }
 
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtEnd()
        {
            // Given
            ListEnd data=new ListEnd(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now,Address="aaa", Learners=new List<Learner>{ new Learner{Name="Lily",Age=20}}},
                       new SessionTime2Elements{ Session=DateTime.Now.AddDays(1),Address="bbb",Learners=new List<Learner>{ new Learner{Name="Joe",Age=19},new Learner{Name="Bruce",Age=20}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{ new Learner{Name="Wuli",Age=28}}}}};
            // When
            using(var designer=new ExportDataDesigner<ListEnd>(data))
            {
            DataTable table=designer.GeneratDataTable();
            // Then
                Assert.Equal("10010",table.Rows[1][1].ToString());
                Assert.Equal("10010",table.Rows[2][1].ToString());
                Assert.Equal("Java",table.Rows[1][0].ToString());
                Assert.Equal("Java",table.Rows[2][0].ToString());
                Assert.Equal("Bill",table.Rows[1][2].ToString());
                Assert.Equal("Bill",table.Rows[2][2].ToString());
                Assert.Equal("aaa",table.Rows[0][6].ToString());
                Assert.Equal("bbb",table.Rows[1][6].ToString());
                Assert.Equal("bbb",table.Rows[2][6].ToString());
                Assert.Equal("ccc",table.Rows[3][6].ToString());
            }

        }    
        [Fact]
        [Trait("Category","MergedCells")]
        public void ValidateMergedCells_AtEnd()
        {
            // Given
            ListEnd data=new ListEnd(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{ new Learner{Name="Lily",Age=20}}},
                       new SessionTime2Elements{ Session=DateTime.Now.AddDays(1),Address="bbb" ,Learners=new List<Learner>{ new Learner{Name="Joe",Age=19},new Learner{Name="Bruce",Age=20}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{ new Learner{Name="Wuli",Age=28}}}}};
            // When
            using(var designer=new ExportDataDesigner<ListEnd>(data))
            {
                DataTable table=designer.GeneratDataTable();
                var mergeCells=designer.MergeCells;
                // Then
                Assert.Equal(5,mergeCells.Count);

                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(0,0,1,4),mergeCells["0-0"]);

                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(1,0,1,4),mergeCells["1-0"]);   
                
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(2,0,1,4),mergeCells["2-0"]);

                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(3,1,1,2),mergeCells["3-1"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int, int, int, int>(6,1,1,2),mergeCells["6-1"]);
            }
        }
        class ListMiddle
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}            
            public List<SessionTime2Elements> Sessions {get;set;}
            public string Trainer {get;set;}
            
        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtMiddle()
        {
            // Given
            ListMiddle data=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="bbb",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}}};
            // When
            using(var designer=new ExportDataDesigner<ListMiddle>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(5,table.Rows.Count);
                Assert.Equal(7,table.Columns.Count);
            }

        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtMiddle()
        {
            // Given
            ListMiddle data=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="bbb",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="ccc",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}}};
           
            // When
            using(var designer=new ExportDataDesigner<ListMiddle>(data))
                {
                DataTable table=designer.GeneratDataTable();
            // Then
                Assert.Equal("10010",table.Rows[1][1].ToString());
                Assert.Equal("10010",table.Rows[4][1].ToString());
                Assert.Equal("Java",table.Rows[1][0].ToString());
                Assert.Equal("Java",table.Rows[4][0].ToString());
                Assert.Equal("Bill",table.Rows[0][6].ToString());
                Assert.Equal("Bill",table.Rows[1][6].ToString());
                Assert.Equal("Bill",table.Rows[2][6].ToString());
                Assert.Equal("Bill",table.Rows[3][6].ToString());
                Assert.Equal("Bill",table.Rows[4][6].ToString());
                Assert.Equal("aaa",table.Rows[0][5].ToString());
                Assert.Equal("aaa",table.Rows[1][5].ToString());
                Assert.Equal("Lily",table.Rows[1][3].ToString());
                Assert.Equal("bbb",table.Rows[2][5].ToString());
                Assert.Equal("bbb",table.Rows[3][5].ToString());
            }

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
            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal(3,table.Rows.Count);
                Assert.Equal(6,table.Columns.Count);
            }

        }
        [Fact]
        [Trait("Category","ExportData Designer")]
        public void ValidateCellValue_DynamicRows_ComprehensiveObj()
        {
            // Given
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Leo",Age=35}}}}                            
                        
                       };         
            // When
            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table=designer.GeneratDataTable();
                // Then
                Assert.Equal("10010",table.Rows[1][1].ToString());
                Assert.Equal("10010",table.Rows[2][1].ToString());
                Assert.Equal("Java",table.Rows[1][0].ToString());
                Assert.Equal("Java",table.Rows[2][0].ToString());
                Assert.Equal("Bill",table.Rows[1][2].ToString());
                Assert.Equal("Bill",table.Rows[2][2].ToString());
                Assert.Equal("Bruce",table.Rows[0][4].ToString());
                Assert.Equal(30,table.Rows[0][5]);
                Assert.Equal("Lily",table.Rows[1][4].ToString());            
                Assert.Equal(20,table.Rows[1][5]);
                Assert.Equal("Leo",table.Rows[2][4].ToString());
                Assert.Equal(35,table.Rows[2][5]);
            }

        }  
        [Fact]
        [Trait("Category","MergedCells")]
        public void GenerateDataTable_DynamicRows_MergeCells()
        {
            // Given
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=30}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Bruce",Age=30}}}}                            
                        
                       };
            // When
            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table=designer.GeneratDataTable();
                Dictionary<string,Tuple<int,int,int,int>> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal<int>(5,mergeCells.Count);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(0,0,1,3),mergeCells["0-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(1,0,1,3),mergeCells["1-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(2,0,1,3),mergeCells["2-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(3,0,1,2),mergeCells["3-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(5,1,1,2),mergeCells["5-1"]);
                
            }

        }
        [Fact]
        [Trait("Category","MergedCells")]
        public void GenerateDataTable_DynamicRows_MergeCells_2()
        {
            // Given
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=30}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Bruce",Age=35},new Learner{Name="Joe",Age=31}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1).AddHours(1),Learners=new List<Learner>{new Learner{Name="Lily",Age=20}}}}                            
                        
                       };
            // When
            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table=designer.GeneratDataTable();
                Dictionary<string,Tuple<int,int,int,int>> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal<int>(5,mergeCells.Count);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(0,0,1,5),mergeCells["0-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(1,0,1,5),mergeCells["1-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(2,0,1,5),mergeCells["2-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(3,0,1,2),mergeCells["3-0"]);
                Assert.Equal<Tuple<int,int,int,int>>(new Tuple<int,int,int,int>(3,2,1,2),mergeCells["3-2"]);

            }

        }

    }
}