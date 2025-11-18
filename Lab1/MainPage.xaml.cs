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
        private bool _showExpressions = false;

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
                RecalculateAll();
                // Запуск логіки RecalculateAll тут гарантує, що всі залежні клітинки оновляться після зміни ExpressionEntry
                if (_activeCell != null && _activeCell.Expression != ExpressionEntry.Text)
                {
                    _activeCell.Expression = ExpressionEntry.Text;
                    RecalculateCell(_activeCell); // Обчислюємо активну клітинку
                }

                 // Переобчислюємо всю таблицю

                // Оновлюємо ExpressionEntry, щоб показати Expression, а не Value, коли вона не активна
                if (_activeCell != null)
                {
                    ExpressionEntry.Text = _activeCell.Expression;
                }

                _activeCell = null;
                if (_activeEntryControl != null)
                {
                    _activeEntryControl.BackgroundColor = Colors.White;
                    _activeEntryControl = null;
                }
            };
        }

        private void DisplayTable()
        {
            TableGrid.RowDefinitions.Clear();
            TableGrid.ColumnDefinitions.Clear();
            TableGrid.Children.Clear();
            _cellEntryControls.Clear();

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
                        Text = _showExpressions ? cell.Expression : (cell.Value?.ToString() ?? cell.Expression),
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
                        if (_activeCell == cell && _activeCell.Expression != entry.Text)
                        {
                            _activeCell.Expression = entry.Text;
                            RecalculateCell(_activeCell);
                            RecalculateAll();
                        }

                        UpdateCellDisplay();
                        entry.BackgroundColor = Colors.White;
                    };

                    entry.TextChanged += (s, e) =>
                    {
                        if (entry.IsFocused)
                        {
                            cell.Expression = entry.Text;
                        }
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

            UpdateCellDisplay();
        }


        private void RecalculateCell(MyCell cell)
        {
            if (string.IsNullOrWhiteSpace(cell.Expression))
            {
                // Очищаємо залежності, якщо вираз порожній
                UpdateDependencies(cell, null);
                cell.SetValue(null); // використовуємо SetValue
                return;
            }

            // Якщо просто число — не парсимо через ANTLR
            if (double.TryParse(cell.Expression, out double number))
            {
                UpdateDependencies(cell, null);
                cell.SetValue(number); // використовуємо SetValue
                return;
            }

            // 1. АНАЛІЗ ВИРАЗУ
            var inputStream = new AntlrInputStream(cell.Expression);
            var lexer = new FormulaLexer(inputStream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new FormulaParser(tokens);

            lexer.RemoveErrorListeners();
            parser.RemoveErrorListeners();

            var parserErrorListener = new ErrorListener();
            parser.AddErrorListener(parserErrorListener);

            var tree = parser.formula();

            // 2. ОБРОБКА СИНТАКСИЧНИХ ПОМИЛОК
            if (parserErrorListener.Errors.Any() || parser.NumberOfSyntaxErrors > 0 || tree == null)
            {
                UpdateDependencies(cell, null);
                cell.SetError("#SYNTAX_ERROR"); // використовуємо SetError
                return;
            }

            // 3. ОНОВЛЕННЯ ГРАФА
            // Тут ми оновлюємо Dependencies та Dependents.
            UpdateDependencies(cell, tree);


            // 4. ПЕРЕВІРКА НА ЦИКЛ
            if (CheckForCycle(cell.Name))
            {
                cell.SetError("#CYCLE!"); // використовуємо SetError
                return;
            }

            // 5. ОБЧИСЛЕННЯ ЗНАЧЕННЯ
            var cellValues = table.Rows
                .SelectMany(r => r.Cells.Values)
                .Where(c => c.Value is double) // Тільки ті, що вже обчислені як double
                .ToDictionary(c => c.Name, c => Convert.ToDouble(c.Value));

            try
            {
                var evaluator = new FormulaEvaluator(cellValues, table);

                double result = evaluator.Visit(tree);

                cell.SetValue(result);
            }
            catch (InvalidOperationException ex) when (ex.Message.StartsWith("Reference error"))
            {
                cell.SetError("#REF!"); // використовуємо SetError
            }
            catch (InvalidOperationException ex) when (ex.Message.StartsWith("Reference to error cell"))
            {
                cell.SetError("#CALC_ERROR"); // використовуємо SetError
            }
            catch (Exception ex)
            {
                cell.SetError("#CALC_ERROR"); // використовуємо SetError
            }
        }

        private void RecalculateAll()
        {
            bool changed;
            const int maxIterations = 3; // Зменшуємо до мінімуму, оскільки цикли виявляються явно
            int iterations = 0;

            var allCells = table.Rows.SelectMany(r => r.Cells.Values).ToList();

            do
            {
                changed = false;
                iterations++;

                foreach (var cell in allCells)
                {
                    // Ми більше не пропускаємо #CYCLE! тут, оскільки може знадобитися його переобчислення,
                    // якщо його вираз був змінений, і CycleCheck() в RecalculateCell це зробить.

                    var oldValue = cell.Value;

                    RecalculateCell(cell);

                    // Перевіряємо, чи змінилося значення
                    if (oldValue == null && cell.Value != null ||
                        oldValue != null && cell.Value == null ||
                        (oldValue != null && cell.Value != null &&
                            // Порівняння значень
                            (oldValue.ToString() != cell.Value.ToString() ||
                            (oldValue is double && cell.Value is double && Math.Abs((double)oldValue - (double)cell.Value) > 1e-9))))
                    {
                        changed = true;
                    }
                }

            } while (changed && iterations < maxIterations);

            // **!!! ВИДАЛЯЄМО ЛОГІКУ ПЕРЕВІРКИ ЦИКЛУ ПІСЛЯ maxIterations !!!**
            // Вона більше не потрібна, оскільки цикли виявляються одразу в RecalculateCell.

            UpdateCellDisplay();

        }

        private void UpdateCellDisplay()
        {
            foreach (var kvp in _cellEntryControls)
            {
                var cellName = kvp.Key;
                var entry = kvp.Value;

                var cell = table.Rows
                    .SelectMany(r => r.Cells.Values)
                    .FirstOrDefault(c => c.Name == cellName);

                if (cell != null)
                {
                    string newText;

                    if (_showExpressions)
                    {
                        // Режим "Вираз": показуємо Expression
                        newText = cell.Expression;
                    }
                    else
                    {
                        // Режим "Значення": показуємо Value або Expression, якщо Value відсутнє/помилкове
                        newText = cell.Error ?? cell.Value?.ToString() ?? cell.Expression;
                    }

                    // Оновлюємо текст, тільки якщо клітинка не є активною (інакше це перерве введення)
                    if (cell != _activeCell)
                    {
                        entry.Text = newText;
                    }
                    else if (cell == _activeCell && !ExpressionEntry.IsFocused)
                    {
                        // Оновлюємо активну клітинку, якщо ExpressionEntry не сфокусована
                        entry.Text = newText;
                    }
                }
            }
        }
        private HashSet<string> UpdateDependencies(MyCell cell, Antlr4.Runtime.Tree.IParseTree tree)
        {
            // 1. Видаляємо всі попередні залежності поточної клітинки
            foreach (var dependentCellName in cell.Dependencies)
            {
                var dependentCell = table.Rows.SelectMany(r => r.Cells.Values).FirstOrDefault(c => c.Name == dependentCellName);
                if (dependentCell != null)
                {
                    dependentCell.Dependents.Remove(cell.Name);
                }
            }
            cell.Dependencies.Clear();

            // 2. Витягуємо нові залежності
            var extractor = new DependencyExtractor();
            if (tree != null)
            {
                extractor.Visit(tree);
            }
            cell.Dependencies = extractor.ReferencedCells;

            // 3. Оновлюємо список "Dependents" у клітинок-донорів
            foreach (var dependentCellName in cell.Dependencies)
            {
                // Тут ми не перевіряємо, чи існує клітинка, оскільки це відбудеться в RecalculateCell.
                // Просто оновлюємо граф, якщо клітинка знайдена в таблиці.
                var dependentCell = table.Rows.SelectMany(r => r.Cells.Values).FirstOrDefault(c => c.Name == dependentCellName);
                if (dependentCell != null)
                {
                    dependentCell.Dependents.Add(cell.Name);
                }
            }

            return cell.Dependencies;
        }
        private bool CheckForCycle(string startCellName)
        {
            var allCells = table.Rows.SelectMany(r => r.Cells.Values).ToDictionary(c => c.Name, c => c);

            // Відвідувані клітинки (для відстеження шляху та уникнення нескінченного циклу)
            var visited = new HashSet<string>();
            // Клітинки на поточному рекурсивному шляху (для виявлення циклу)
            var recursionStack = new HashSet<string>();

            return DFS(startCellName, allCells, visited, recursionStack);
        }

        private bool DFS(string cellName, Dictionary<string, MyCell> allCells, HashSet<string> visited, HashSet<string> recursionStack)
        {
            if (!allCells.TryGetValue(cellName, out var cell))
            {
                // Якщо клітинка не існує, це не цикл (буде #REF!), але зупиняємо пошук
                return false;
            }

            if (recursionStack.Contains(cellName))
            {
                return true; // ЗНАЙДЕНО ЦИКЛ!
            }

            if (visited.Contains(cellName))
            {
                return false; // Клітинка вже відвідана і не є частиною циклу
            }

            visited.Add(cellName);
            recursionStack.Add(cellName);

            foreach (var dependencyName in cell.Dependencies)
            {
                if (DFS(dependencyName, allCells, visited, recursionStack))
                {
                    return true;
                }
            }

            recursionStack.Remove(cellName);
            return false;
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

        private void ToggleModeButton_Clicked(object sender, EventArgs e)
        {
            // 1. Інвертуємо прапорець
            _showExpressions = !_showExpressions;

            // 2. Оновлюємо текст кнопки для зворотного зв'язку
            if (sender is Button button)
            {
                button.Text = _showExpressions ? "Показати Значення" : "Показати Вирази";
            }

            // 3. Оновлюємо відображення всієї таблиці
            UpdateCellDisplay();
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            table.AddRow();
            DisplayTable();
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (table.RowCount == 1)
            {
                _ = DisplayAlert("Помилка", "В таблиці має бути принаймні один рядок.", "OK");
                return;
            }

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
            // Prevent deleting when table is already 1x1: show informative alert
            if ( table.ColCount == 1)
            {
                _ = DisplayAlert("Помилка", "В таблиці має бути принаймні один стовпчик.", "OK");
                return;
            }

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
