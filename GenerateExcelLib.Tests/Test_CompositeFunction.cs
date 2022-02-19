using Xunit;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using GenerateExcelLib;

namespace GenerateExcelLib.Tests
{
    public class Test_CompositeFunction
    {
        class Learner
        {
            public string Name{get;set;}
            public int Age {get;set;}
        }
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
                    //using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
                    using MemoryStream ms=new MemoryStream();
                    ExportRegularExcel work_book=new ExportRegularExcel(ms);
                    DrawParameter parameter=new DrawParameter{
                        StartRow=1,StartCol=1,
                        MergeCells=designer.MergeCells,
                        HiddenColumns=designer.HiddenCols
                    };
                    // When run test function

                        work_book.DrawExcel(mydata,parameter,true);
                        work_book.Save();
                        //Then Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(7,result.Item1); //assert column num
                        Assert.Equal(4,result.Item2); //assert row num

                        
                    
                }
            }
            
        } 
        [Fact]
        [Trait("Category","Assemble")]
        public void Export_ComplexContentObject_from5Row_WithNoHead()
        {
            // Given

            using(var designer=new ExportDataDesigner<ComprehensiveObj>(data))
            {
            
                //generate datatable
                using(DataTable mydata=designer.GeneratDataTable())
                {
                    using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
                    //using MemoryStream ms=new MemoryStream();
                    ExportRegularExcel work_book=new ExportRegularExcel(ms);
                    DrawParameter parameter=new DrawParameter{
                        StartRow=5,StartCol=1,
                        MergeCells=designer.MergeCells,
                        HiddenColumns=designer.HiddenCols
                    };
                    // When run test function

                    work_book.DrawExcel(mydata,parameter,false);
                    work_book.Save();
                        //Then Assert result
                        var result=Excel_Ops_Aspose.Retrieve_Num_Column_Row(ms);
                        Assert.Equal(7,result.Item1); //assert column num
                        Assert.Equal(7,result.Item2); //assert row num
                    
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
                    //using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
                    using MemoryStream ms=new MemoryStream();
                    ExportRegularExcel work_book=new ExportRegularExcel(ms);
                    DrawParameter parameter=new DrawParameter{
                        StartRow=1,StartCol=1,
                        MergeCells=designer.MergeCells,
                        HiddenColumns=designer.HiddenCols
                    };

                        //When: run test function
                        work_book.DrawExcel(mydata,parameter,true);
                        work_book.Save();
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
                    //using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
                    using MemoryStream ms=new MemoryStream();
                    ExportRegularExcel work_book=new ExportRegularExcel(ms);
                    DrawParameter parameter=new DrawParameter{
                        StartRow=1,StartCol=1,
                        MergeCells=designer.MergeCells,
                        HiddenColumns=designer.HiddenCols
                    };                      
                      //When: run test function
                        work_book.DrawExcel(mydata,parameter,true);
                        work_book.Save();                   

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
                    //using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
                    using MemoryStream ms=new MemoryStream();
                    ExportRegularExcel work_book=new ExportRegularExcel(ms);
                    DrawParameter parameter=new DrawParameter{
                        StartRow=1,StartCol=1,
                        MergeCells=designer.MergeCells,
                        HiddenColumns=designer.HiddenCols
                    };
                    
                        work_book.DrawExcel(mydata,parameter,false);
                        work_book.Save();
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

        [Fact]
        [Trait("Category", "Test Class")]
        public void MergeCell_ComplexContentObject_TestClass()
        {
            #region // Given
            var data = new ScheduleClass
            {
                Coordinator = "Facilitator 1: Role 1\nFacilitator 2: Role 2",
                ClassTitle = "Class1",
                TimeSlots = new List<TimeSlots>
                {
                    new TimeSlots
                    {
                        Date="Mon, 10 Jan 2022",
                        DateTime = "Mon, 10 Jan 2022\n(09:00 AM - 10:00 AM)",
                        SessionName = "Session1",
                        Modality = "F2F",
                        Facilitator = "Trainer 1: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge 1",
                                GoogleMapLink = ""
                            },
                            new Venues
                            {
                                Venue = "Learning Lounge 2",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 1",
                            DateEstimatedTimeofArrival = "Mon, 10 Jan 2022 / 9:00 AM",
                            Designation = "GD (Grassroots)",
                            VehicleDetails = "SJP 170 P / Bronze Subaru SUV"
                        }
                    },
                    new TimeSlots
                    {
                        Date="Mon, 10 Jan 2022",
                        DateTime = "Mon, 10 Jan 2022\n(10:00 AM - 01:00 PM)",
                        SessionName = "Session2",
                        Modality = "",
                        Facilitator = "Trainer 1: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge 1",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 1\nStakeholder 2\nStakeholder 3",
                            DateEstimatedTimeofArrival = "Mon, 10 Jan 2022 / 9:00 AM\n\nMon, 10 Jan 2022 / 9:30 AM",
                            Designation = "GD (Grassroots)\nACE (Operations)\nCEO",
                            VehicleDetails = "SJP 170 P / Bronze Subaru SUV\nSGA 9007 J / Toyota Sedan / White\n"
                        }
                    },
                    new TimeSlots
                    {
                        Date="Mon, 10 Jan 2022",
                        DateTime = "Mon, 10 Jan 2022\n(02:00 PM - 06:00 PM)",
                        SessionName = "Session3",
                        Modality = "",
                        Facilitator = "Trainer 1: Role 1\nTrainer 2: Role 2",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge 3",
                                GoogleMapLink = ""
                            },
                            new Venues
                            {
                                Venue = "Learning Lounge 4",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "",
                            DateEstimatedTimeofArrival = "",
                            Designation = "",
                            VehicleDetails = ""
                        }
                    },
                    new TimeSlots
                    {
                        Date="Mon, 10 Jan 2022",
                        DateTime = "Mon, 10 Jan 2022\n(07:00 PM - 09:00 PM)",
                        SessionName = "Session1",
                        Modality = "",
                        Facilitator = "Trainer 1: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            },
                            new Venues
                            {
                                Venue = "Learning Lounge 1",
                                GoogleMapLink = ""
                            }                            
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 1",
                            DateEstimatedTimeofArrival = "Mon, 10 Jan 2022 / 9:00 AM",
                            Designation = "GD (Grassroots)",
                            VehicleDetails = "SJP 170 P / Bronze Subaru SUV"
                        }
                    },
                    new TimeSlots
                    {
                        Date="Tue, 11 Jan 2022",
                        DateTime = "Tue, 11 Jan 2022\n(09:00 AM - 01:00 PM)",
                        SessionName = "Session5",
                        Modality = "",
                        Facilitator = "Trainer 2: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            },
                            new Venues
                            {
                                Venue = "Learning Lounge 1",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 3",
                            DateEstimatedTimeofArrival = "Tue, 11 Jan 2022 / 10:00 AM",
                            Designation = "",
                            VehicleDetails = ""
                        }
                    },
                    new TimeSlots
                    {
                        Date="Tue, 11 Jan 2022",
                        DateTime = "Tue, 11 Jan 2022\n(02:00 PM - 04:00 PM)",
                        SessionName = "Session6",
                        Modality = "",
                        Facilitator = "Trainer 3: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "",
                            DateEstimatedTimeofArrival = "",
                            Designation = "",
                            VehicleDetails = ""
                        }
                    },
                    new TimeSlots
                    {
                        Date="Tue, 11 Jan 2022",
                        DateTime = "Tue, 11 Jan 2022\n(04:00 PM - 06:00 PM)",
                        SessionName = "Session7",
                        Modality = "F2F",
                        Facilitator = "Trainer 4: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = "",
                            },
                            new Venues
                            {
                                Venue = "Learning Lounge 5",
                                GoogleMapLink = "",
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 1",
                            DateEstimatedTimeofArrival = "Tue, 11 Jan 2022 / 10:00 AM",
                            Designation = "CEO",
                            VehicleDetails = "SJP 170 P / Bronze Subaru SUV"
                        }
                    },
                    new TimeSlots
                    {
                        Date="Tue, 11 Jan 2022",
                        DateTime = "Tue, 11 Jan 2022\n(07:00 PM - 09:00 PM)",
                        SessionName = "Session7",
                        Modality = "F2F",
                        Facilitator = "Trainer 4: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Zone 4\nBLK 616\nLevel 3 Room 2",
                                GoogleMapLink = "https://www.google.com/maps/place/"
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "Stakeholder 1",
                            DateEstimatedTimeofArrival = "Tue, 11 Jan 2022 / 10:00 AM",
                            Designation = "CEO",
                            VehicleDetails = "SJP 170 P / Bronze Subaru SUV"
                        }
                    },
                    new TimeSlots
                    {
                        Date="Wed, 12 Jan 2022",
                        DateTime = "Wed, 12 Jan 2022\n(09:00 AM - 01:00 PM)",
                        SessionName = "Session8",
                        Modality = "F2F",
                        Facilitator = "External Trainer",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "",
                        NoofPax = "",
                        VDincharge = "",
                        StakeHolders = new StakeHolders
                        {
                            NameofStakeholder = "",
                            DateEstimatedTimeofArrival = "",
                            Designation = "",
                            VehicleDetails = ""
                        }
                    }
                },
                ClassCode = "C-22-0067"
            };
            #endregion
            using var designer = new ExportDataDesigner<ScheduleClass>(data);

            //Given: generate datatable
            using DataTable mydata = designer.GeneratDataTable();
            //using FileStream ms=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
            using MemoryStream ms=new MemoryStream();
            ExportRegularExcel work_book=new ExportRegularExcel(ms);
            DrawParameter parameter=new DrawParameter{
                StartRow=1,StartCol=1,
                MergeCells=designer.MergeCells,
                HiddenColumns=designer.HiddenCols
            };


            work_book.DrawExcel(mydata,parameter,true);
            work_book.Save();
            Assert.Equal(50, designer.MergeCells.Count);
            // Assert.Equal(new Tuple<int, int, int, int>(3, 6, 1, 2), designer.MergeCells["3-6"]);
            // Assert.Equal(new Tuple<int, int, int, int>(4, 6, 1, 2), designer.MergeCells["4-6"]);
            // Assert.Equal(new Tuple<int, int, int, int>(5, 6, 1, 2), designer.MergeCells["5-6"]);
            // Assert.Equal(new Tuple<int, int, int, int>(6, 0, 1, 7), designer.MergeCells["6-0"]);
            // Assert.Equal(new Tuple<int, int, int, int>(8, 0, 1, 9), designer.MergeCells["8-0"]);
            // Assert.Equal(new Tuple<int, int, int, int>(15, 0, 1, 9), designer.MergeCells["15-0"]);
        }

    }
    public class ScheduleClass
    {
        public string Coordinator { get; set; }
        public string ClassTitle { get; set; }
        public List<TimeSlots> TimeSlots { get; set; }

        public string ClassCode { get; set; }
    }

    public class TimeSlots
    {
        [MergeIdentifier("Venues")]
        public string Date{get;set;}
        public string DateTime { get; set; }
        [MergeIdentifier("Session")]
        public string SessionName { get; set; }
        [MergeFollower("Session")]
        public string Modality { get; set; }
        public string Facilitator { get; set; }
        public List<Venues> Venues { get; set; }
        public string ForWhom { get; set; }
        public string NoofPax { get; set; }
        public string VDincharge { get; set; }
        public StakeHolders StakeHolders { get; set; }
    }
    public class Venues
    {
        [MergeFollower("Venues")]
        public string Venue { get; set; }
        public string GoogleMapLink { get; set; }
    }
    public class StakeHolders
    {
        public string NameofStakeholder { get; set; }
        public string DateEstimatedTimeofArrival { get; set; }
        public string Designation { get; set; }
        public string VehicleDetails { get; set; }
    }


}