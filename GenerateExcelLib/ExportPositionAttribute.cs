using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateExcelLib
{
    /// <summary>
    /// mark this att on export property set order when export to excel
    /// TODO only allow add attribute on property
    /// </summary>
    public class ExportPositionAttribute : Attribute
    {
        public string ColName { get; set; }

        public ExportPositionAttribute(string ColName)
        {
            this.ColName = ColName;
        }

          
    }
}
