using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace GenerateExcelLib.Tests
{
    public class Test_MergeCell : IDisposable
    {
        [Fact]
        [Trait("Category", "MergeCell")]
        public void Init_MergeCellObject_With_Negative_StartRow()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new MergeCell(-1, 1, 1, 1));
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Init_MergeCellObject_With_Negative_StartColumn()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new MergeCell(1, -1, 1, 1));
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Init_MergeCellObject_With_Negative_TotalRows()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new MergeCell(1, 1, -1, 1));
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Init_MergeCellObject_With_Negative_TotalColumns()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new MergeCell(1, 1, 1, -1));
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Add_OffSet_One_Row_One_Column()
        {
            MergeCell mergeCell = new MergeCell(1, 1, 1, 1);
            mergeCell.AddOffSet(1, 1);
            Assert.Equal(2, mergeCell.StartRow);
            Assert.Equal(2, mergeCell.StartColumn);
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Add_OffSet_MinusOne_Row_One_Column()
        {
            MergeCell mergeCell = new MergeCell(1, 1, 1, 1);
            mergeCell.AddOffSet(-1, 1);
            Assert.Equal(0, mergeCell.StartRow);
            Assert.Equal(2, mergeCell.StartColumn);
        }

        [Fact]
        [Trait("Category", "MergeCell")]
        public void Add_OffSet_MinusOne_And_Throw_OutOfRangeException()
        {
            MergeCell mergeCell = new MergeCell(0, 1, 1, 1);
            Assert.Throws<ArgumentOutOfRangeException>(() => mergeCell.AddOffSet(-1, 0));
        }

        public void Dispose()
        {
            // release resource if you use them during test.
        }
    }
}
