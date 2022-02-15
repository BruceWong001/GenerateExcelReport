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
    public class ExportAttrAttribute : Attribute
    {
        public string ColName { get; set; }

        public ExportType ExportType { get; set; }

        public ExportAttrAttribute(string ColName, ExportType exportType)
        {
            this.ColName = ColName;
            this.ExportType = exportType;
        }
    }

    public enum ExportType
    {
        Plaintext = 0,
        Hyperlink = 1
    }

    public enum ExportExtendedKey
    {
        ColumnName = 0,
        ColumnType = 1,
        CanCombine = 2
    }
}
