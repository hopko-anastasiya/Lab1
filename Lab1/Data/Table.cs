using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace Lab1.Data
{
    public class Table
    {
        public List<Row> Rows { get; set; } = new List<Row>();
        public int ColCount { get; private set; }
        public int RowCount { get; private set; }

        public Table(int rows, int cols)
        {
            if (rows <= 0 || cols <= 0)
            {
                throw new ArgumentException("Кількість рядків та стовпців має бути більше нуля.");
            }

            RowCount = rows;
            ColCount = cols;

            for (int r = 0; r < RowCount; r++)
            {
                var newRow = new Row(r, ColCount);
                Rows.Add(newRow);
            }
        }

        public void AddRow()
        {
            var newRow = new Row(RowCount, ColCount);
            Rows.Add(newRow);
            RowCount++;
        }

        public void DeleteRow()
        {
            if (RowCount > 1)
            {
                Rows.RemoveAt(RowCount - 1);
                RowCount--;
            }
            
        }

        public void AddColumn()
        {
            int newColIndex = ColCount;

            for (int r = 0; r < RowCount; r++)
            {
                var row = Rows[r];
                string cellName = CellName.GetName(newColIndex, r);

                row.Cells.Add(newColIndex, new Cell { Name = cellName });
            }

            ColCount++;
        }

        public void DeleteColumn()
        {
            if (ColCount > 1)
            {
                int lastColIndex = ColCount - 1;

                for (int r = 0; r < RowCount; r++)
                {
                    var row = Rows[r];
                    row.Cells.Remove(lastColIndex);
                }

                ColCount--;
            }
        }

        public Cell this[int rowIndex, int columnIndex]
        {
            get
            {
                if (rowIndex >= 0 && rowIndex < RowCount && columnIndex >= 0 && columnIndex < ColCount)
                {
                    return Rows[rowIndex].Cells[columnIndex];
                }
                throw new IndexOutOfRangeException("Індекси клітинки знаходяться за межами таблиці.");
            }
        }

        public Cell GetCellByName(string name)
        {
            try
            {
                var (col, row) = CellName.ParseName(name);

                if (row < 0 || row >= RowCount || col < 0 || col >= ColCount)
                {
                    throw new IndexOutOfRangeException($"Клітинка '{name}' знаходиться за межами таблиці.");
                }

                // Рядок має індекс 0-based
                if (!Rows[row].Cells.TryGetValue(col, out var cell))
                {
                    // Це повинно статися лише, якщо таблиця була створена некоректно
                    throw new Exception($"Не вдалося знайти клітинку '{name}' за координатами.");
                }

                return cell;
            }
            catch (FormatException ex)
            {
                throw new ArgumentException($"Некоректний формат імені клітинки '{name}'.", ex);
            }
            catch (IndexOutOfRangeException ex)
            {
                throw new ArgumentException($"Клітинка '{name}' не існує.", ex);
            }
            catch (Exception ex)
            {
                // Обробляємо інші потенційні помилки, пов'язані з доступом до списку/словника
                throw new InvalidOperationException($"Помилка при отриманні клітинки '{name}'.", ex);
            }
        }

    }
}
