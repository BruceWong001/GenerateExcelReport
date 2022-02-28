# GenerateExcelReport

I wanna write a common lib to generate Excel for report purpose.   
**The Example like below picture:**   
![](/assets/images/ExcelReport.png)  

So far, it uses Aspose and ClosedXML to draw Excel, you can choose one of them to use. If you have the license of Aspose, you can use Aspose by default, since it's the most stable and effective solution by now.   
Please reference by below quick start. If you have any idea, please let us know.  

### Quick Start  
1. **Define your owned Data**  
First you should define you data which will be converted to DataTable, please take a look below the code snip. The Generic List it will generate multiple rows by the number of itmes. Other Non-Genric List Property will be automatically merged, since the value will be same.
The Attribute "MergeIdentifier"and"MergeFollower" is used by pairing. you should set same Key name like "Session", it means the merge rule of MergeFollower won't only be merged by self-value, but also same "SessionName".    
There is a boolean parameter for "MergeIdentifier", that means if you want to delete this column after merge cell, it's mostly used when some column only for merge rule but it won't display for the end user.
```c#
    public class CustmerData
    {
        public string Date{get;set;}
        public string DateTime { get; set; }
        [MergeIdentifier("Session",false)]
        public string SessionName { get; set; }
        [MergeFollower("Session")]
        public string Modality { get; set; }
        public string Facilitator { get; set; }
        public List<Venues> Venues { get; set; }
        public string Whom { get; set; }
        public StakeHolders StakeHolders { get; set; }
    }

```  
2. **Your frist sample.**  
After generate you defined data, you can easily to use the below code to quickly generate a excel.  
```c#
    //create designer by customer Data 
    using var designer = new ExportDataDesigner<CustmerData>(data);
    //conver customer data to data table object 
    using DataTable mydata = designer.GeneratDataTable();
    //create file stream for excel file
    using FileStream fs=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
    // create closedXML provider with file stream object.
    var work_book=new ExportRegularExcelClosedXML(fs);
    // draw a excel file by data table
    work_book.DrawExcel(mydata);
    work_book.Save();    

```  
3. **More parameter as configuration**
You can create a DrawParameter object to pass more information to the lib, so that you can draw excel for your special needs.  
+ You can speicify the position to draw table in excel: Startrow, StartCol  
+ If you don't wanna merge cell, you can let MergeCells as null.  
+ you wannt remove some of column after draw the Excel.  

```c#

    //create designer by customer Data 
    using var designer = new ExportDataDesigner<CustmerData>(data);
    //conver customer data to data table object 
    using DataTable mydata = designer.GeneratDataTable();
    //create file stream for excel file
    using FileStream fs=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
    // create closedXML provider with file stream object.
    var work_book=new ExportRegularExcelClosedXML(fs);
    // using below object for multiple config parameter for draw excel func
    DrawParameter parameter=new DrawParameter{
        StartRow=5,StartCol=3,
        MergeCells=designer.MergeCells,
        HiddenColumns=designer.HiddenCols
    };
    // draw a excel file by data table false mean no header for the table
    work_book.DrawExcel(mydata,parameter,false);
    work_book.Save();

```  
4. All above is for draw one data table, you also can draw multiple data table and save them finally. It could be contolled by your business logic.  

``` c# 

    //conver customer data to data table object 
    using DataTable mydata = designer.GeneratDataTable();
    //create file stream for excel file
    using FileStream fs=new FileStream(@"c:\testnew.xlsx",FileMode.Create);
    // create closedXML provider with file stream object.
    var work_book=new ExportRegularExcelClosedXML(fs);
    // if there is one more 
    for(var CustomerData in List<CustomerData>)
    {
        //create designer by customer Data 
        using var designer = new ExportDataDesigner<CustmerData>(CustomerData);

        // draw a excel file by data table
        work_book.DrawExcel(mydata);

    }
    work_book.Save();

```  

It welcomes if you want to contribute any code, idea or feedback.
