using Lab1.Data;

using System.IO;
using CommunityToolkit.Maui.Storage;
using ClosedXML.Excel;
using System.Numerics;
using Microsoft.Maui.Controls;

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


            ExpressionEntry.TextChanged += (s, e) =>
            {
                if (_activeCell != null)
                {
                    _activeCell.Expression = ExpressionEntry.Text;
                }
            };

            ExpressionEntry.Completed += (s, e) =>
            {
                if (_activeCell != null)
                {
                    // Примусово обчислюємо та оновлюємо
                    _activeCell.Expression = ExpressionEntry.Text;

                    // Втрачаємо фокус, щоб спрацювала логіка RecalculateAll
                    ExpressionEntry.Unfocus();
                }
            };

            DisplayTable();

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
                    Text = ((char)('A' + c)).ToString(),
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

                        ExpressionEntry.Text = cell.Expression;
                        entry.BackgroundColor = Color.FromArgb("#502A3756");
                    };

                    entry.Unfocused += (s, e) =>
                    {
                        if (_activeEntryControl == entry) 
                            entry.BackgroundColor = Color.FromArgb("#502A3756"); 
                        else
                            entry.BackgroundColor = Colors.White;
                    };
                    
                    entry.TextChanged += (s, e) => { 
                        if (_activeCell == cell) 
                            cell.Expression = entry.Text;
                    };

                    TableGrid.Add(border, c + 1, r + 1);
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
