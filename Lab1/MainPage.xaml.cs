using Lab1.Data;

using System.IO;
using CommunityToolkit.Maui.Storage;
using ClosedXML.Excel;
using System.Numerics;
using Microsoft.Maui.Controls;
using Antlr4.Runtime;

using MyCell = Lab1.Data.Cell;

namespace Lab1
{
    public partial class MainPage : ContentPage
    {
        private Table table;

        private Lab1.Data.Cell? _activeCell;
        private Entry? _activeEntryControl;

        private readonly Dictionary<string, Entry> _cellEntryControls = new(StringComparer.OrdinalIgnoreCase);


        public MainPage()
        {
            InitializeComponent();

            table = new Table(10, 10);

            DisplayTable();

            ExpressionEntry.Unfocused += (s, e) =>
            {
                // Запуск логіки RecalculateAll тут гарантує, що всі залежні клітинки оновляться після зміни ExpressionEntry
                if (_activeCell != null && _activeCell.Expression != ExpressionEntry.Text)
                {
                    _activeCell.Expression = ExpressionEntry.Text;
                    RecalculateCell(_activeCell); // Обчислюємо активну клітинку
                }

                RecalculateAll(); // Переобчислюємо всю таблицю

                // Оновлюємо ExpressionEntry, щоб показати Expression, а не Value, коли вона не активна
                if (_activeCell != null)
                {
                    ExpressionEntry.Text = _activeCell.Expression;
                }

                if (_activeEntryControl != null)
                {
                    _activeEntryControl.Text = _activeCell?.Value?.ToString() ?? _activeCell?.Expression ?? "";
                    _activeEntryControl.BackgroundColor = Colors.White; // Гарантуємо, що колір повернеться до білого
                }

                _activeCell = null;
                _activeEntryControl = null;
            };
        }

        private void DisplayTable()
        {
            TableGrid.RowDefinitions.Clear();
            TableGrid.ColumnDefinitions.Clear();
            TableGrid.Children.Clear();

            int rows = table.RowCount;
            int cols = table.ColCount;

            TableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = Microsoft.Maui.GridLength.Auto });
            for (int c = 0; c < cols; c++)
                TableGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = Microsoft.Maui.GridLength.Auto });

            TableGrid.RowDefinitions.Add(new RowDefinition { Height = Microsoft.Maui.GridLength.Auto });
            for (int r = 0; r < rows; r++)
                TableGrid.RowDefinitions.Add(new RowDefinition { Height = Microsoft.Maui.GridLength.Auto });

            for (int c = 0; c < cols; c++)
            {
                var label = new Label
                {
                    Text = CellName.GetColumnName(c),
                    FontAttributes = FontAttributes.Bold,
                    HorizontalTextAlignment = Microsoft.Maui.TextAlignment.Center,
                    VerticalTextAlignment = Microsoft.Maui.TextAlignment.Center,
                    BackgroundColor = Color.FromArgb("#2A3756"),
                    TextColor = Colors.White,
                    Padding = 4
                };
                var border = new Border
                {
                    Content = label,
                    Stroke = Colors.Black,
                    StrokeThickness = 0.2,
                    Padding = 0,
                    Margin = 0
                };
                TableGrid.Add(border, c + 1, 0);
            }

            for (int r = 0; r < rows; r++)
            {
                var label = new Label
                {
                    Text = (r + 1).ToString(),
                    FontAttributes = FontAttributes.Bold,
                    HorizontalTextAlignment = Microsoft.Maui.TextAlignment.Center,
                    VerticalTextAlignment = Microsoft.Maui.TextAlignment.Center,
                    BackgroundColor = Color.FromArgb("#2A3756"),
                    TextColor = Colors.White,
                    Padding = 4
                };
                var border = new Border
                {
                    Content = label,
                    Stroke = Colors.Black,
                    StrokeThickness = 0.1,
                };
                TableGrid.Add(border, 0, r + 1);
            }

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cell = table.Rows[r].Cells[c];

                    string displayValue = cell.Value?.ToString() ?? cell.Expression;

                    var entry = new Entry
                    {
                        Text = displayValue,
                        BackgroundColor = Colors.White,
                        HorizontalTextAlignment = Microsoft.Maui.TextAlignment.Center,
                        WidthRequest = 70,
                        TextColor = Color.FromArgb("#2A3756")
                    };

                    var border = new Border
                    {
                        Content = entry,
                        Stroke = Colors.Black,
                        StrokeThickness = 0.1,
                    };

                    _cellEntryControls[cell.Name] = entry;

                    entry.Focused += (s, e) =>
                    {
                        if (_activeEntryControl != null)
                            _activeEntryControl.BackgroundColor = Colors.White;

                        _activeCell = cell;
                        _activeEntryControl = entry;

                        // Показуємо Expression у верхньому полі
                        ExpressionEntry.Text = cell.Expression;

                        // Встановлюємо фокус на ExpressionEntry
                        ExpressionEntry.Focus();

                        entry.BackgroundColor = Color.FromArgb("#502A3756");
                    };

                    entry.Unfocused += (s, e) =>
                    {
                        // Якщо Unfocus відбувся на клітинці, але вона не була активною для ExpressionEntry
                        if (_activeCell == cell && _activeCell.Expression != entry.Text)
                        {
                            _activeCell.Expression = entry.Text;
                            RecalculateCell(_activeCell);
                            RecalculateAll();
                        }

                        // Оновлюємо відображення Entry після обчислення
                        entry.Text = cell.Value?.ToString() ?? cell.Expression;
                        entry.BackgroundColor = Colors.White;

                        // НЕ СКИДАЄМО _activeCell ТУТ, це робить ExpressionEntry.Unfocused
                    };

                    entry.TextChanged += (s, e) =>
                    {
                        // Оновлюємо Expression клітинки. Це відбувається перед Recalculate
                        cell.Expression = entry.Text;
                    };


                    TableGrid.Add(border, c + 1, r + 1);
                }
            }


            ExpressionEntry.TextChanged += (s, e) =>
            {
                if (_activeCell != null && ExpressionEntry.IsFocused)
                {
                    _activeCell.Expression = ExpressionEntry.Text;
                }
            };

            ExpressionEntry.Completed += (s, e) =>
            {
                // Втрачаємо фокус, що запускає логіку в ExpressionEntry.Unfocused
                ExpressionEntry.Unfocus();
            };
        }


        private void RecalculateCell(MyCell cell)
        {
            if (string.IsNullOrWhiteSpace(cell.Expression))
            {
                cell.Value = null;
                return;
            }

            // Якщо просто число — не парсимо через ANTLR
            if (double.TryParse(cell.Expression, out double number))
            {
                cell.Value = number;
                return;
            }

            var cellValues = table.Rows
                .SelectMany(r => r.Cells.Values)
                .Where(c => c.Value is double) // Тільки ті, що вже обчислені як double
                .ToDictionary(c => c.Name, c => Convert.ToDouble(c.Value));

            // Скидаємо Value, щоб уникнути використання старого значення у випадку помилки
            object? oldValue = cell.Value;
            cell.Value = null;

            try
            {
                var inputStream = new AntlrInputStream(cell.Expression);
                var lexer = new FormulaLexer(inputStream);
                var tokens = new CommonTokenStream(lexer);
                var parser = new FormulaParser(tokens);

                // ВИПРАВЛЕННЯ: Додаємо обробку помилок
                lexer.RemoveErrorListeners();
                parser.RemoveErrorListeners();

                var parserErrorListener = new ErrorListener();
                parser.AddErrorListener(parserErrorListener);

                var tree = parser.formula();

                // Якщо є синтаксичні помилки
                if (parserErrorListener.Errors.Any() || parser.NumberOfSyntaxErrors > 0 || tree == null)
                {
                    cell.Value = "#SYNTAX_ERROR";
                    return;
                }

                var evaluator = new FormulaEvaluator(cellValues);

                cell.Value = evaluator.Visit(tree);
            }
            catch (Exception ex)
            {
                // Обробка помилок обчислення (наприклад, ділення на нуль, неіснуюча клітинка)
                // Можемо залишити Value як текст помилки
                cell.Value = "#CALC_ERROR";
            }
        }

        private void RecalculateAll()
        {
            bool changed;
            const int maxIterations = 50; // Обмеження для запобігання нескінченному циклу (може вказувати на циклічну залежність)
            int iterations = 0;

            do
            {
                changed = false;
                iterations++;

                // Створюємо список усіх клітинок для ітерації
                var allCells = table.Rows.SelectMany(r => r.Cells.Values).ToList();

                foreach (var cell in allCells)
                {
                    var oldValue = cell.Value;

                    RecalculateCell(cell);

                    // Перевіряємо, чи змінилося значення
                    if (oldValue == null && cell.Value != null ||
                        oldValue != null && cell.Value == null ||
                        (oldValue != null && cell.Value != null &&
                            // Порівняння значень, враховуючи, що Value може бути рядком помилки або числом
                            (oldValue.ToString() != cell.Value.ToString() ||
                            (oldValue is double && cell.Value is double && Math.Abs((double)oldValue - (double)cell.Value) > 1e-9))))
                    {
                        changed = true;
                    }
                }

            } while (changed && iterations < maxIterations); // Повторюємо, доки є зміни або не досягнено ліміту ітерацій

            // Оновлюємо відображення в Entry (цей цикл залишається без змін)
            // ... (ваш існуючий код оновлення відображення)
            foreach (var kvp in _cellEntryControls)
            {
                var cellName = kvp.Key;
                var entry = kvp.Value;

                var cell = table.Rows
                    .SelectMany(r => r.Cells.Values)
                    .FirstOrDefault(c => c.Name == cellName);

                // ВИПРАВЛЕННЯ: Якщо клітинка активна, не змінюємо її текст, щоб не переривати введення
                if (cell != null && cell != _activeCell)
                    entry.Text = cell.Value?.ToString() ?? cell.Expression;
                else if (cell != null && cell == _activeCell && !ExpressionEntry.IsFocused)
                {
                    // Оновлюємо активну клітинку, якщо ExpressionEntry не сфокусована
                    entry.Text = cell.Value?.ToString() ?? cell.Expression;
                }
            }


        }

        private async Task SaveFile()
        {
            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Таблиця");

                for (int r = 0; r < table.RowCount; r++)
                {
                    for (int c = 0; c < table.ColCount; c++)
                    {
                        var cellExpression = table[r, c].Expression;

                        if (!string.IsNullOrEmpty(cellExpression))
                        {
                            worksheet.Cell(r + 1, c + 1).Value = cellExpression;
                        }
                    }
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;

                var result = await FileSaver.Default.SaveAsync("table_data.xlsx", stream, CancellationToken.None);

                if (result.IsSuccessful)
                {
                    await DisplayAlert("Успіх", $"Файл успішно збережено.\nШлях: {result.FilePath}", "OK");
                }
                else
                {
                    await DisplayAlert("Скасовано", "Збереження файлу скасовано або не вдалося.", "OK");
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка збереження", $"Не вдалося зберегти файл: {ex.Message}", "OK");
            }
        }
        private async Task OpenFile()
        {
            try
            {
                var customFileType = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
            {
                { DevicePlatform.Android, new[] { ".xlsx" } },
                { DevicePlatform.iOS, new[] { ".xlsx" } },
                { DevicePlatform.WinUI, new[] { ".xlsx" } },
                { DevicePlatform.macOS, new[] { ".xlsx" } }
            });

                var options = new PickOptions
                {
                    FileTypes = customFileType,
                    PickerTitle = "Виберіть файл Excel (.xlsx)"
                };

                var result = await FilePicker.Default.PickAsync(options);

                if (result == null) return;

                using var stream = await result.OpenReadAsync();
                using var workbook = new XLWorkbook(stream);

                var worksheet = workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    await DisplayAlert("Помилка", "Файл Excel не містить робочих листів.", "OK");
                    return;
                }

                int maxRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                int maxColUsed = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                const int minExpectedSize = 10;

                int rows = Math.Max(maxRowUsed, minExpectedSize);
                int cols = Math.Max(maxColUsed, minExpectedSize);

                if (rows == 0 || cols == 0)
                {
                    await DisplayAlert("Помилка", "Робочий лист порожній.", "OK");
                    return;
                }

                var newTable = new Table(rows, cols);

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var excelCell = worksheet.Cell(r + 1, c + 1);

                        string expression = excelCell.FormulaA1;

                        if (string.IsNullOrEmpty(expression))
                        {
                            expression = excelCell.Value.ToString();
                        }

                        newTable[r, c].Expression = expression;
                    }
                }

                table = newTable;
                DisplayTable();

                _activeCell = null;
                if (_activeEntryControl != null)
                {
                    _activeEntryControl.BackgroundColor = Colors.White;
                    _activeEntryControl = null;
                }
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка завантаження", $"Не вдалося завантажити файл Excel: {ex.Message}", "OK");
            }
        }
        private async Task NewFile()
        {
            bool saveNeeded = await DisplayAlert("Створення нової таблиці", "Створити нову таблицю? Незбережені дані будуть втрачені.\n\nЗберегти поточну таблицю?", "Так, зберегти", "Ні, не зберігати");

            if (saveNeeded)
            {
                await SaveFile();
            }

            table = new Table(10, 10);
            DisplayTable();
        }

        

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            table.AddRow();
            DisplayTable();
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            table.DeleteRow();
            DisplayTable();
        }
        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            table.AddColumn();
            DisplayTable();
        }
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            table.DeleteColumn();
            DisplayTable();
        }
        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            // Просто викликаємо асинхронну логіку
            await SaveFile();
        }
        private async void OpenButton_Clicked(object sender, EventArgs e)
        {

            await OpenFile();
        }
        private async void NewButton_Clicked(object sender, EventArgs e)
        {
            // Просто викликаємо асинхронну логіку
            await NewFile();
        }
        private async void AboutButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Вітаємо в табличному редакторі! \n \nДоступні операції та функції: \n1) +, -, *, / (бінарні операції) \n2) +, - (унарні операції) \n3) ^ (піднесення у степінь) \n4) =, <, > \n5) <=, >=, <> \n \nПриємного вам користування:)", "OK");
        }
        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно бажаєте вийти?", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }
    }

}
