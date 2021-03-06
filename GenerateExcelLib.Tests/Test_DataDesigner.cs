using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;

namespace GenerateExcelLib.Tests
{    

    public class Test_ExportDataDesigner
    {
        class Learner
        {
            public string Name{get;set;}
            public int Age {get;set;}
        }
        class SessionTime
        {
            public DateTime Session{get;set;}
            [MergeIdentifier("SessionName",true)]
            public string SessionName{get;set;}
        }
        class SessionTime2Elements
        {
            public DateTime Session{get;set;}
            public List<Learner> Learners{get;set;}
            public string Address{get;set;}
        }        
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
                Assert.Equal(new MergeCell(0,2,2,1),mergeCells["2-0"]);

                Assert.Equal(new MergeCell(0,3,2,1),mergeCells["3-0"]);
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

                Assert.Equal(new MergeCell(0,0,4,1),mergeCells["0-0"]);

                Assert.Equal(new MergeCell(0,1,4,1),mergeCells["1-0"]);   
                
                Assert.Equal(new MergeCell(0,2,4,1),mergeCells["2-0"]);

                Assert.Equal(new MergeCell(1,3,2,1),mergeCells["3-1"]);
                Assert.Equal(new MergeCell(1,6,2,1),mergeCells["6-1"]);
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
        [Fact]
        [Trait("Category","MergedCells")]
        public void ValidateMergeCellCount_AtMiddle_OverlapRegion()
        {
            // Given
            ListMiddle data=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="aaa",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=18}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="aaa",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}
                       }};
           
            // When
            using(var designer=new ExportDataDesigner<ListMiddle>(data))
            {
                DataTable table=designer.GeneratDataTable();
            // Then
                Assert.Equal<int>(6, designer.MergeCells.Count);
            }

        }  
        [Fact]
        [Trait("Category","MergedCells")]
        public void ValidateMergeCellContent_AtMiddle_OverlapRegion()
        {
            // Given
            ListMiddle data=new ListMiddle(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       Sessions= new List<SessionTime2Elements>{new SessionTime2Elements{Session=DateTime.Now, Address="aaa", Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=19}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(1),Address="aaa",Learners=new List<Learner>{new Learner{Name="Joe",Age=29},new Learner{Name="Andy",Age=18}}},
                       new SessionTime2Elements{Session=DateTime.Now.AddDays(2),Address="aaa",Learners=new List<Learner>{new Learner{Name="Nancy",Age=20}}}
                       }};
           
            // When
            using(var designer=new ExportDataDesigner<ListMiddle>(data))
            {
                DataTable table=designer.GeneratDataTable();
                var mergeCells=designer.MergeCells;
            // Then
                Assert.Equal<int>(6, mergeCells.Count);                
                Assert.Equal(new MergeCell(0,0,5,1),mergeCells["0-0"]);
                Assert.Equal(new MergeCell(0,1,5,1),mergeCells["1-0"]);
                Assert.Equal(new MergeCell(0,2,2,1),mergeCells["2-0"]);
                Assert.Equal(new MergeCell(2,2,2,1),mergeCells["2-2"]);
                Assert.Equal(new MergeCell(0,5,5,1),mergeCells["5-0"]);
                Assert.Equal(new MergeCell(0,6,5,1),mergeCells["6-0"]);
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
                Dictionary<string, MergeCell> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal<int>(5,mergeCells.Count);
                Assert.Equal(new (0,0,3,1),mergeCells["0-0"]);
                Assert.Equal(new (0,1,3,1),mergeCells["1-0"]);
                Assert.Equal(new (0,2,3,1),mergeCells["2-0"]);
                Assert.Equal(new (0,3,2,1),mergeCells["3-0"]);
                Assert.Equal(new (1,5,2,1),mergeCells["5-1"]);
                
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
                Dictionary<string, MergeCell> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal<int>(5,mergeCells.Count);
                Assert.Equal(new MergeCell(0,0,5,1),mergeCells["0-0"]);
                Assert.Equal(new MergeCell(0,1,5,1),mergeCells["1-0"]);
                Assert.Equal(new MergeCell(0,2,5,1),mergeCells["2-0"]);
                Assert.Equal(new MergeCell(0,3,2,1),mergeCells["3-0"]);
                Assert.Equal(new MergeCell(2,3,2,1),mergeCells["3-2"]);

            }

        }
        class ClassInfo
        {
            public string ClassCode{get;set;}
            public string ClassName{get;set;}
            public List<SessionTime> SessionInfo{get;set;}
            [MergeFollower("SessionName")]
            public string Venue{get;set;}
        }
        [Fact]
        [Trait("Category","MergedCells")]
        public void GenerateDataTable_MergeCells_withIdentifierRule()
        {
            // Given
            ClassInfo data=new ClassInfo(){ClassName="Java",ClassCode="10010", Venue="Building No1",
                       
                       SessionInfo=new List<SessionTime>{new SessionTime{Session=DateTime.Now,SessionName="Sub 1"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="Sub 1"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="Sub 2"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1).AddHours(1),SessionName="Sub 2"}}                            
                        
                       };
            // When
            using(var designer=new ExportDataDesigner<ClassInfo>(data))
            {
                DataTable table=designer.GeneratDataTable();
                Dictionary<string, MergeCell> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal<int>(7,mergeCells.Count);
                Assert.Equal(new MergeCell(0,0,4,1),mergeCells["0-0"]);
                Assert.Equal(new MergeCell(0,1,4,1),mergeCells["1-0"]);
                Assert.Equal(new MergeCell(1,2,2,1),mergeCells["2-1"]);
                Assert.Equal(new MergeCell(0,3,2,1),mergeCells["3-0"]);
                Assert.Equal(new MergeCell(2,3,2,1),mergeCells["3-2"]);
                Assert.Equal(new MergeCell(0,4,2,1),mergeCells["4-0"]);
                Assert.Equal(new MergeCell(2,4,2,1),mergeCells["4-2"]);
            }

        }
        [Fact]
        [Trait("Category","DeleteCells")]
        public void GenerateDataTable_Return_DelColumn()
        {
            // Given
            ClassInfo data=new ClassInfo(){ClassName="Java",ClassCode="10010", Venue="Building No1",
                       
                       SessionInfo=new List<SessionTime>{new SessionTime{Session=DateTime.Now,SessionName="Sub 1"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="Sub 1"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1),SessionName="Sub 2"},
                                            new SessionTime{Session=DateTime.Now.AddDays(1).AddHours(1),SessionName="Sub 2"}}                            
                        
                       };
            // When
            using(var designer=new ExportDataDesigner<ClassInfo>(data))
            {
                DataTable table=designer.GeneratDataTable();
                Dictionary<string, MergeCell> mergeCells=designer.MergeCells;
                // Then
                Assert.Equal(3,designer.HiddenCols[0]);
                Assert.Equal(1,designer.HiddenCols.Count);
            }

        }


    }
}