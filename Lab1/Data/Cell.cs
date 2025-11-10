using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab1.Data
{
    public enum CellStatus
    {
        Dirty,
        ParsingError,
        CalculationError,
        Calculating,
        Ready
    }
    public class Cell
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public object? Value { get; set; }

        public void SetValue(object value)
        {
            this.Value = value;
        }
    }
}
