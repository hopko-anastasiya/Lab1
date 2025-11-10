using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab1.Data
{
    public static class CellName
    {
        public static string GetColumnName(int col)
        {
            if (col < 0) return string.Empty;

            // Використовуємо індексацію від 0
            int dividend = col + 1;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        // Отримує повне ім'я клітинки (наприклад, "A1" для (0, 0))
        public static string GetName(int col, int row)
        {
            // Рядки індексуються від 1, стовпці - від 0
            return GetColumnName(col) + (row + 1).ToString();
        }

        // Парсить ім'я клітинки (наприклад, "B5") у координати (col=1, row=4)
        public static (int col, int row) ParseName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Ім'я клітинки не може бути порожнім.");

            string letters = "";
            string numbers = "";

            foreach (char c in name.ToUpperInvariant())
            {
                if (char.IsLetter(c))
                    letters += c;
                else if (char.IsDigit(c))
                    numbers += c;
                else
                    throw new FormatException($"Некоректний символ в імені клітинки: {name}");
            }

            if (string.IsNullOrWhiteSpace(letters) || string.IsNullOrWhiteSpace(numbers))
                throw new FormatException($"Некоректний формат імені клітинки: {name}");

            // Парсимо номер рядка (віднімаємо 1, бо в таблиці індексація від 0)
            if (!int.TryParse(numbers, out int rowNumber) || rowNumber <= 0)
                throw new FormatException($"Некоректний номер рядка в імені клітинки: {name}");

            int row = rowNumber - 1;

            // Парсимо індекс стовпця (A=0, B=1...)
            int col = 0;
            for (int i = 0; i < letters.Length; i++)
            {
                col *= 26;
                col += (letters[i] - 'A' + 1);
            }
            col--; // Повертаємо індексацію до 0

            return (col, row);
        }
    }
}
