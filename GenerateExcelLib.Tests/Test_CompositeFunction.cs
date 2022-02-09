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
        [Trait("Category","Assemble")]
        public void Export_ComplexContentObject_WithHead()
        {
            // Given
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Leo",Age=35}}}}                            
                        
                       };         

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
                        Assert.Equal(6,result.Item1); //assert column num
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
            ComprehensiveObj data=new ComprehensiveObj(){ClassTitle="Java",ClassCode="10010",Trainer="Bill",
                       
                       SessionList=new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Leo",Age=35}}}}                            
                        
                       };         

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
                        Assert.Equal(6,result.Item1); //assert column num
                        Assert.Equal(7,result.Item2); //assert row num

                        }
                    
                }
            }
            
        } 
    }


}