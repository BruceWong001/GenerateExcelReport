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
        public string Name { get; set; }
        public int Age { get; set; }
    }
    class SessionTime
    {
        public DateTime Session { get; set; }
    }

    public class Test_ExportDataDesigner
    {
        class SimpleClass
        {
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
            public string Trainer { get; set; }
        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_SimpleObj()
        {
            // Given
            SimpleClass data = new SimpleClass() { ClassTitle = "Java", ClassCode = "10010", Trainer = "Bill" };
            // When
            using (var designer = new ExportDataDesigner<SimpleClass>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(1, table.Rows.Count);
                Assert.Equal(3, table.Columns.Count);
                Assert.Equal("Bill", table.Rows[0][2].ToString());
            }

        }

        class SimpleClassEx
        {
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
            public Learner Student { get; set; }
            public string Trainer { get; set; }
            public DateTime RegistryTime { get; set; }
        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_SimpleObj_WithObjMember()
        {
            // Given
            SimpleClassEx data = new SimpleClassEx() { ClassTitle = "Java", ClassCode = "10010", Trainer = "Bill", Student = new Learner { Name = "Bruce", Age = 30 }, RegistryTime = DateTime.Now };
            // When
            using (var designer = new ExportDataDesigner<SimpleClassEx>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(1, table.Rows.Count);
                Assert.Equal(6, table.Columns.Count);
                Assert.Equal("Bill", table.Rows[0][4].ToString());
            }

        }
        class ListStart
        {
            public List<SessionTime> Sessions { get; set; }
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtBegin()
        {
            // Given
            ListStart data = new ListStart()
            {
                Sessions = new List<SessionTime>() { new SessionTime { Session = DateTime.Now }, new SessionTime { Session = DateTime.Now.AddDays(1) } },
                ClassTitle = "Java",
                ClassCode = "10010"
            };
            // When
            using (var designer = new ExportDataDesigner<ListStart>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(3, table.Columns.Count);
            }
        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtBegin()
        {
            // Given
            ListStart data = new ListStart()
            {
                Sessions = new List<SessionTime>() { new SessionTime { Session = DateTime.Now }, new SessionTime { Session = DateTime.Now.AddDays(1) } },
                ClassTitle = "Java",
                ClassCode = "10010"
            };
            // When
            using (var designer = new ExportDataDesigner<ListStart>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal("10010", table.Rows[0][2].ToString());
                Assert.Equal("10010", table.Rows[1][2].ToString());
                Assert.Equal("Java", table.Rows[0][1].ToString());
                Assert.Equal("Java", table.Rows[1][1].ToString());
            }
        }
        [Fact]
        [Trait("Category", "MergedCells")]
        public void ValidateMergedCells_AtBegin()
        {
            // Given
            ListStart data = new ListStart()
            {
                Sessions = new List<SessionTime>() { new SessionTime { Session = DateTime.Now }, new SessionTime { Session = DateTime.Now.AddDays(1) } },
                ClassTitle = "Java",
                ClassCode = "10010"
            };
            // When
            using (var designer = new ExportDataDesigner<ListStart>(data))
            {
                DataTable table = designer.GeneratDataTable();
                var mergeCells = designer.MergeCells;
                // Then
                Assert.Equal(2, mergeCells.Count);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(1, 0, 1, 2), mergeCells["1-0"]);

                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(2, 0, 1, 2), mergeCells["2-0"]);
            }
        }
        class ListEnd
        {
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
            public string Trainer { get; set; }
            public List<Learner> learners { get; set; }
        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtEnd()
        {
            // Given
            ListEnd data = new ListEnd()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",
                learners = new List<Learner> { new Learner { Name = "Lily", Age = 20 }, new Learner { Name = "Joe", Age = 19 }, new Learner { Name = "Wuli", Age = 28 } }
            };
            // When
            using (var designer = new ExportDataDesigner<ListEnd>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(3, table.Rows.Count);
                Assert.Equal(5, table.Columns.Count);
            }

        }

        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtEnd()
        {
            // Given
            ListEnd data = new ListEnd()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",
                learners = new List<Learner> { new Learner { Name = "Lily", Age = 20 }, new Learner { Name = "Joe", Age = 19 }, new Learner { Name = "Wuli", Age = 28 } }
            };
            // When
            using (var designer = new ExportDataDesigner<ListEnd>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal("10010", table.Rows[1][1].ToString());
                Assert.Equal("10010", table.Rows[2][1].ToString());
                Assert.Equal("Java", table.Rows[1][0].ToString());
                Assert.Equal("Java", table.Rows[2][0].ToString());
                Assert.Equal("Bill", table.Rows[1][2].ToString());
                Assert.Equal("Bill", table.Rows[2][2].ToString());
            }

        }
        [Fact]
        [Trait("Category", "MergedCells")]
        public void ValidateMergedCells_AtEnd()
        {
            // Given
            ListEnd data = new ListEnd()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",
                learners = new List<Learner> { new Learner { Name = "Lily", Age = 20 }, new Learner { Name = "Joe", Age = 20 }, new Learner { Name = "Wuli", Age = 28 } }
            };
            // When
            using (var designer = new ExportDataDesigner<ListEnd>(data))
            {
                DataTable table = designer.GeneratDataTable();
                var mergeCells = designer.MergeCells;
                // Then
                Assert.Equal(4, mergeCells.Count);

                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(0, 0, 1, 3), mergeCells["0-0"]);

                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(1, 0, 1, 3), mergeCells["1-0"]);

                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(2, 0, 1, 3), mergeCells["2-0"]);

                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(4, 0, 1, 2), mergeCells["4-0"]);

            }
        }
        class ListMiddle
        {
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
            public List<SessionTime> Sessions { get; set; }
            public string Trainer { get; set; }

        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_DynamicRows_AtMiddle()
        {
            // Given
            ListMiddle data = new ListMiddle()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",
                Sessions = new List<SessionTime> { new SessionTime { Session = DateTime.Now }, new SessionTime { Session = DateTime.Now.AddDays(1) }, new SessionTime { Session = DateTime.Now.AddDays(2) } }
            };
            // When
            using (var designer = new ExportDataDesigner<ListMiddle>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(3, table.Rows.Count);
                Assert.Equal(4, table.Columns.Count);
            }

        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void ValidateCellValue_DynamicRows_AtMiddle()
        {
            // Given
            ListMiddle data = new ListMiddle()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",
                Sessions = new List<SessionTime> { new SessionTime { Session = DateTime.Now }, new SessionTime { Session = DateTime.Now.AddDays(1) }, new SessionTime { Session = DateTime.Now.AddDays(2) } }
            };
            // When
            using (var designer = new ExportDataDesigner<ListMiddle>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal("10010", table.Rows[1][1].ToString());
                Assert.Equal("10010", table.Rows[2][1].ToString());
                Assert.Equal("Java", table.Rows[1][0].ToString());
                Assert.Equal("Java", table.Rows[2][0].ToString());
                Assert.Equal("Bill", table.Rows[1][3].ToString());
                Assert.Equal("Bill", table.Rows[2][3].ToString());
            }

        }
        class SessionObj
        {
            public DateTime Session { get; set; }
            public List<Learner> Learners { get; set; }
        }
        class ComprehensiveObj
        {
            public string ClassTitle { get; set; }
            public string ClassCode { get; set; }
            public string Trainer { get; set; }
            public List<SessionObj> SessionList { get; set; }

        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void GenerateDataTable_DynamicRows_Validate_ColRow_Num()
        {
            // Given
            ComprehensiveObj data = new ComprehensiveObj()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",

                SessionList = new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Bruce",Age=30}}}}

            };
            // When
            using (var designer = new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal(3, table.Rows.Count);
                Assert.Equal(6, table.Columns.Count);
            }

        }
        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void ValidateCellValue_DynamicRows_ComprehensiveObj()
        {
            // Given
            ComprehensiveObj data = new ComprehensiveObj()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",

                SessionList = new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=30},new Learner{Name="Lily",Age=20}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Leo",Age=35}}}}

            };
            // When
            using (var designer = new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table = designer.GeneratDataTable();
                // Then
                Assert.Equal("10010", table.Rows[1][1].ToString());
                Assert.Equal("10010", table.Rows[2][1].ToString());
                Assert.Equal("Java", table.Rows[1][0].ToString());
                Assert.Equal("Java", table.Rows[2][0].ToString());
                Assert.Equal("Bill", table.Rows[1][2].ToString());
                Assert.Equal("Bill", table.Rows[2][2].ToString());
                Assert.Equal("Bruce", table.Rows[0][4].ToString());
                Assert.Equal(30, table.Rows[0][5]);
                Assert.Equal("Lily", table.Rows[1][4].ToString());
                Assert.Equal(20, table.Rows[1][5]);
                Assert.Equal("Leo", table.Rows[2][4].ToString());
                Assert.Equal(35, table.Rows[2][5]);
            }

        }
        [Fact]
        [Trait("Category", "MergedCells")]
        public void GenerateDataTable_DynamicRows_MergeCells()
        {
            // Given
            ComprehensiveObj data = new ComprehensiveObj()
            {
                ClassTitle = "Java",
                ClassCode = "10010",
                Trainer = "Bill",

                SessionList = new List<SessionObj>{new SessionObj{Session=DateTime.Now,Learners=new List<Learner>{new Learner{Name="Bruce",Age=20},new Learner{Name="Lily",Age=30}}},
                                            new SessionObj{Session=DateTime.Now.AddDays(1),Learners=new List<Learner>{new Learner{Name="Bruce",Age=30}}}}

            };
            // When
            using (var designer = new ExportDataDesigner<ComprehensiveObj>(data))
            {
                DataTable table = designer.GeneratDataTable();
                Dictionary<string, Tuple<int, int, int, int>> mergeCells = designer.MergeCells;
                // Then
                Assert.Equal<int>(5, mergeCells.Count);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(0, 0, 1, 3), mergeCells["0-0"]);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(1, 0, 1, 3), mergeCells["1-0"]);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(2, 0, 1, 3), mergeCells["2-0"]);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(3, 0, 1, 2), mergeCells["3-0"]);
                Assert.Equal<Tuple<int, int, int, int>>(new Tuple<int, int, int, int>(5, 1, 1, 2), mergeCells["5-1"]);

            }

        }

        public class ExportSchool
        {
            public ExportSchool()
            {
                Students = new List<ExportStudent>();
            }

            [ExportAttr("A", ExportType.Plaintext)]
            public string SchoolName { get; set; }

            public List<ExportStudent> Students { get; set; }
        }

        public class ExportStudent
        {
            [ExportAttr("B", ExportType.Plaintext)]
            public int Age { get; set; }

            [ExportAttr("C", ExportType.Plaintext)]
            public string Name { get; set; }
        }

        [Fact]
        [Trait("Category", "ExportData Designer")]
        public void ValidateDateTable_Caption()
        {
            var data = new ExportSchool()
            {
                SchoolName = "AvePoint",
                Students = new List<ExportStudent>()
                {
                    new ExportStudent()
                    {
                        Age = 18,
                        Name = "Evans"
                    }
                }
            };
            using (var designer = new ExportDataDesigner<ExportSchool>(data))
            {
                DataTable table = designer.GeneratDataTable();
                Assert.Equal("A", table.Columns[0].ExtendedProperties[ExportExtendedKey.ColumnName]);
                Assert.Equal("B", table.Columns[1].ExtendedProperties[ExportExtendedKey.ColumnName]);
                Assert.Equal("C", table.Columns[2].ExtendedProperties[ExportExtendedKey.ColumnName]);
                Assert.Equal(ExportType.Plaintext, table.Columns[0].ExtendedProperties[ExportExtendedKey.ColumnType]);
            }
        }

        public class ScheduleClass
        {
            public string Coordinator { get; set; }
            public string ClassTitle { get; set; }
            public List<TimeSlots> TimeSlots { get; set; }

            public string ClassCode { get; set; }
        }

        [Fact]
        [Trait("Category", "Test Class")]
        public void MergeCell_ComplexContentObject_TestClass()
        {
            // Given
            var data = new ScheduleClass
            {
                Coordinator = "Facilitator 1: Role 1\nFacilitator 2: Role 2",
                ClassTitle = "Class1",
                TimeSlots = new List<TimeSlots>
                {
                    new TimeSlots
                    {
                        DateTime = "Mon, 10 Jan 2022\n(09:00 AM - 10:00 AM)",
                        SessionName = "Session1",
                        Modality = "F2F",
                        Facilitator = "Trainer 1: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                        DateTime = "Mon, 10 Jan 2022\n(10:00 AM - 01:00 PM)",
                        SessionName = "Session2",
                        Modality = "",
                        Facilitator = "Trainer 1: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                        DateTime = "Mon, 10 Jan 2022\n(02:00 PM - 06:00 PM)",
                        SessionName = "Session3",
                        Modality = "",
                        Facilitator = "Trainer 1: Role 1\nTrainer 2: Role 2",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Learning Lounge",
                                GoogleMapLink = ""
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                        ForWhom = "PA Staff\nGRL",
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
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                        DateTime = "Tue, 11 Jan 2022\n(07:00 PM - 09:00 PM)",
                        SessionName = "Session7",
                        Modality = "F2F",
                        Facilitator = "Trainer 4: Role 1",
                        Venues = new List<Venues>
                        {
                            new Venues
                            {
                                Venue = "Zone 4\nBLK 616\nLevel 3 Room 2",
                                GoogleMapLink = "https://www.google.com/maps/place/singapore+poscode"
                            }
                        },
                        ForWhom = "PA Staff\nGRL",
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
                        ForWhom = "PA Staff\nGRL",
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
            using var designer = new ExportDataDesigner<ScheduleClass>(data);

            //Given: generate datatable
            using DataTable mydata = designer.GeneratDataTable();
            var work_book = new ExportRegularExcel();

            using var Result_Book = work_book.GenerateExcel(mydata);
            //When: run test function
            work_book.MergeCell(Result_Book, designer.MergeCells);

            using MemoryStream ms = new(new byte[5000000]);
            //save excel file content into tempfile(memory stream)
            Result_Book.Save(ms, SaveFormat.Xlsx);
            Result_Book.Save(@"c:\test3.xlsx"); // only for debug
                                               // first col (one based),first row (one based), total cols(one based), total rows(one based)
                                               //Then: Assert result
            Assert.Equal(19, designer.MergeCells.Count);
            //Assert.Equal(new Tuple<int, int, int, int>(3, 6, 1, 2), designer.MergeCells["3-6"]);
            //Assert.Equal(new Tuple<int, int, int, int>(4, 6, 1, 2), designer.MergeCells["4-6"]);
            //Assert.Equal(new Tuple<int, int, int, int>(5, 6, 1, 2), designer.MergeCells["5-6"]);
            //Assert.Equal(new Tuple<int, int, int, int>(6, 0, 1, 4), designer.MergeCells["6-0"]);
            //Assert.Equal(new Tuple<int, int, int, int>(6, 4, 1, 2), designer.MergeCells["6-4"]);
            //Assert.Equal(new Tuple<int, int, int, int>(8, 0, 1, 9), designer.MergeCells["8-0"]);
            //Assert.Equal(new Tuple<int, int, int, int>(15, 0, 1, 9), designer.MergeCells["15-0"]);
        }

        public class TimeSlots
        {
            public string DateTime { get; set; }
            public string SessionName { get; set; }
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
}