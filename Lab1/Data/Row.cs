using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab1.Data
{
    public class Row
    {
        public int Index { get; set; }
        public Dictionary<int, Cell> Cells { get; set; }

        public Row(int index, int columnCount)
        {
            Index = index;
            Cells = new Dictionary<int, Cell>();
            for (int col = 0; col < columnCount; col++)
            {
                string cellName = CellName.GetName(col, index);
                Cells.Add(col, new Cell { Name = cellName });
            }

        }
    }
}
