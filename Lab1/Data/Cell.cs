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

        public string? Error { get; set; }

        public HashSet<string> Dependencies { get; set; } = new HashSet<string>();
        public HashSet<string> Dependents { get; set; } = new HashSet<string>();

        public void SetValue(object? value)
        {
            this.Value = value;
            this.Error = null;
        }

        public void SetError(string errorMessage)
        {
            this.Value = null; // Помилка обчислення скидає значення
            this.Error = errorMessage;
        }
        public string DisplayText
        {
            get
            {
                if (!string.IsNullOrEmpty(Error))
                    return Error;

                return Value?.ToString() ?? Expression;
            }
        }
    }
}
