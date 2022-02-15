using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;
using Aspose.Cells;


namespace GenerateExcelLib.Tests
{
    public class Test_CompositeFunction
    {
        class SessionObj
        {
            public DateTime Session {get;set;}
            public string Teacher{get;set;}
            public List<Learner> Learners {get;set;}
        }
        class ComprehensiveObj
        {   
            public string ClassTitle{get;set;}
            public string ClassCode{get;set;}            
            public string Trainer {get;set;}
            public List<SessionObj> SessionList{get;set;}
            
        }
        private ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Teacher="Linda",Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Teacher="Lucy",Learners=new List<Learner>{new Learner{Name="Leo",Age=35}}}}};     
        [Fact]
        [Trait("Category","Assemble")]
        public void Export_ComplexContentObject_WithHead()
        {
            // Given    

            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
            
                //generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    var work_book=new ExportRegularExcel();
                    // When run test function
                    var Result_Book= work_book.GenerateExcel(mydata);
    
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        Result_Book.Save(ms,SaveFormat.Xlsx);
                       // Result_Book.Save(@"c:\test.xlsx"); // only for debug
                        //Then Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(7,result.Item1); //assert column num
                        Assert.Equal(4,result.Item2); //assert row num

                        }
                    
                }
            }
            
        } 
        [Fact]
        [Trait("Category","Assemble")]
        public void Export_ComplexContentObject_WithNoHead()
        {
            // Given

            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
            
                //generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    var work_book=new ExportRegularExcel();
                    // When run test function
                    var Result_Book= work_book.GenerateExcel(mydata,5,false);
    
                    using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                    {
                        //save excel file content into tempfile(memory stream)
                        Result_Book.Save(ms,SaveFormat.Xlsx);
                    // Result_Book.Save(@"c:\test.xlsx"); // only for debug
                        //Then Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(7,result.Item1); //assert column num
                        Assert.Equal(7,result.Item2); //assert row num

                        }
                    
                }
            }
            
        } 
        [Fact]
        [Trait("Category","Assemble")]
        public void MergeCell_ComplexContentObject_withHead()
        {
            // Given     

            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
            
                //Given: generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    var work_book=new ExportRegularExcel();
                    
                    using(var Result_Book= work_book.GenerateExcel(mydata))
                    {
                        //When: run test function
                        work_book.MergeCell(Result_Book,designer.MergeCells); 
                    
                        using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                        {
                            //save excel file content into tempfile(memory stream)
                            Result_Book.Save(ms,SaveFormat.Xlsx);
                          //  Result_Book.Save(@"c:\test.xlsx"); // only for debug
                          // first col (one based),first row (one based), total cols(one based), total rows(one based)
                            //Then: Assert result
                            Assert.Equal<int>(5,designer.MergeCells.Count);
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,1,2,1,3)); //assert the spicified area is merged.
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,2,1,3));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,3,2,1,3));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,4,2,1,2));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,5,2,1,2));
                        }
                    
                    }
                }
            }
            
        } 
        [Fact]
        [Trait("Category","Assemble")]
        public void MergeCell_ComplexContentObject_withHead_2()
        {
            // Given    
            ComprehensiveObj data_onlyThisfunc=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Teacher="Linda",Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Teacher="Lucy",Learners=new List<Learner>{new Learner{Name="Leo",Age=25},new Learner{Name="Joe",Age=20}}}}};    

            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data_onlyThisfunc))
            {
            
                //Given: generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    var work_book=new ExportRegularExcel();
                    
                    using(var Result_Book= work_book.GenerateExcel(mydata))
                    {
                        //When: run test function
                        work_book.MergeCell(Result_Book,designer.MergeCells); 
                    
                        using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                        {
                            //save excel file content into tempfile(memory stream)
                            Result_Book.Save(ms,SaveFormat.Xlsx);
                          //  Result_Book.Save(@"c:\test.xlsx"); // only for debug
                          // first col (one based),first row (one based), total cols(one based), total rows(one based)
                            //Then: Assert result
                            Assert.Equal<int>(7,designer.MergeCells.Count);
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,1,2,1,4)); //assert the spicified area is merged.
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,2,1,4));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,3,2,1,4));
                           Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,4,2,1,2));
                           Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,4,4,1,2));
                           Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,5,2,1,2));
                           Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,5,4,1,2));
                        }
                    
                    }
                }
            }
            
        } 
        [Fact]
        [Trait("Category","Assemble")]
        public void MergeCell_ComplexContentObject_withNoHead()
        {
            // Given    
            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
            
                //Given: generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    var work_book=new ExportRegularExcel();
                    
                    using(var Result_Book= work_book.GenerateExcel(mydata,1,false))
                    {
                        //When: run test function
                        work_book.MergeCell(Result_Book,designer.MergeCells,false); 
                    
                        using(MemoryStream ms=new MemoryStream(new byte[5000000]))
                        {
                            //save excel file content into tempfile(memory stream)
                            Result_Book.Save(ms,SaveFormat.Xlsx);
                           // Result_Book.Save(@"c:\test.xlsx"); // only for debug
                          // first col (one based),first row (one based), total cols(one based), total rows(one based)
                            //Then: Assert result
                            Assert.Equal<int>(5,designer.MergeCells.Count);
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,1,1,1,3)); //assert the spicified area is merged.
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,2,1,1,3));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,3,1,1,3));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,4,1,1,2));
                            Assert.True(Excel_Ops_Aspose.Is_MergeCell(ms,5,1,1,2));
                        }
                    
                    }
                }
            }
            
        } 
    }


}