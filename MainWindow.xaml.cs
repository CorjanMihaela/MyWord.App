using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;

namespace MyWord.App;

public partial class MainWindow : Window
{
    private string? _currentFile;
    private bool _isDirty;
    private bool _suppressChanges;
    private bool _isDarkTheme;

    private const double AcademicFontSize = 12;
    private const double AcademicLineHeight = 18;
    private const double FirstLineIndent = 47.24;

    private const string Filter =
        "Rich Text (*.rtf)|*.rtf|Text (*.txt)|*.txt";

    /*
     * HeadingLevel:
     * 0 = text normal
     * 1 = capitol
     * 2 = subcapitol
     */
    public static readonly DependencyProperty HeadingLevelProperty =
        DependencyProperty.RegisterAttached(
            "HeadingLevel",
            typeof(int),
            typeof(MainWindow),
            new PropertyMetadata(0));

    public static void SetHeadingLevel(
        DependencyObject element,
        int value)
    {
        element.SetValue(HeadingLevelProperty, value);
    }

    public static int GetHeadingLevel(
        DependencyObject element)
    {
        return (int)element.GetValue(HeadingLevelProperty);
    }

    public MainWindow()
    {
        InitializeComponent();

        ConfigureFonts();
        ConfigureSizes();

        ApplyAcademicDocumentFormat();
        ApplyLightTheme();

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Gata");
    }

    private void ConfigureFonts()
    {
        FontBox.ItemsSource = Fonts.SystemFontFamilies
            .OrderBy(font => font.Source)
            .ToList();

        FontBox.SelectedItem =
            Fonts.SystemFontFamilies
                .FirstOrDefault(font => font.Source == "Times New Roman")
            ?? Fonts.SystemFontFamilies.FirstOrDefault();
    }

    private void ConfigureSizes()
    {
        SizeBox.ItemsSource = new double[]
        {
            10,
            11,
            12,
            13,
            14,
            16,
            18,
            20,
            24,
            28,
            36
        };

        SizeBox.SelectedItem = AcademicFontSize;
    }

    private void Editor_TextChanged(
        object sender,
        TextChangedEventArgs e)
    {
        if (_suppressChanges)
            return;

        _isDirty = true;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Modificat");
    }

    private void Editor_PreviewKeyDown(
        object sender,
        System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;

        var paragraph = GetCurrentParagraph();

        if (paragraph is null)
            return;

        var headingLevel = GetHeadingLevel(paragraph);

        if (headingLevel == 0)
            return;

        /*
         * Dacă utilizatorul apasă Enter după un capitol sau subcapitol,
         * următorul paragraf devine automat text normal.
         */
        e.Handled = true;

        var normalParagraph = CreateNormalParagraph();

        Editor.Document.Blocks.InsertAfter(
            paragraph,
            normalParagraph);

        PlaceCaretAtEnd(normalParagraph);

        _isDirty = true;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Text normal");
    }

    private void New_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (!ConfirmDiscardChanges())
            return;

        try
        {
            _suppressChanges = true;

            var document = new FlowDocument
            {
                PagePadding = new Thickness(20)
            };

            document.Blocks.Add(
                CreateNormalParagraph());

            Editor.Document = document;
        }
        finally
        {
            _suppressChanges = false;
        }

        _currentFile = null;
        _isDirty = false;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Document nou");

        Editor.Focus();
    }

    private void Open_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (!ConfirmDiscardChanges())
            return;

        var dialog = new OpenFileDialog
        {
            Filter = Filter,
            Title = "Deschide document"
        };

        if (dialog.ShowDialog() != true)
            return;

        try
        {
            _suppressChanges = true;

            var extension = Path
                .GetExtension(dialog.FileName)
                .ToLowerInvariant();

            var dataFormat = extension == ".rtf"
                ? DataFormats.Rtf
                : DataFormats.Text;

            var range = new TextRange(
                Editor.Document.ContentStart,
                Editor.Document.ContentEnd);

            using var fileStream = File.OpenRead(dialog.FileName);

            range.Load(fileStream, dataFormat);

            RestoreHeadingMetadata();

            _currentFile = dialog.FileName;
            _isDirty = false;

            UpdateTitle();
            UpdateStatistics();
            UpdateStatus("Document deschis");

            Editor.Focus();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Documentul nu a putut fi deschis.\n\n{ex.Message}",
                "Eroare la deschidere",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            _suppressChanges = false;
        }
    }

    private void Save_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentFile))
        {
            SaveAs_Click(sender, e);
            return;
        }

        SaveDocument();
    }

    private void SaveAs_Click(
        object sender,
        RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = Filter,
            DefaultExt = ".rtf",
            AddExtension = true,
            Title = "Salvează document ca..."
        };

        if (dialog.ShowDialog() != true)
            return;

        _currentFile = dialog.FileName;

        SaveDocument();
    }

    private void SaveDocument()
    {
        if (string.IsNullOrWhiteSpace(_currentFile))
            return;

        try
        {
            var extension = Path
                .GetExtension(_currentFile)
                .ToLowerInvariant();

            var dataFormat = extension == ".rtf"
                ? DataFormats.Rtf
                : DataFormats.Text;

            var range = new TextRange(
                Editor.Document.ContentStart,
                Editor.Document.ContentEnd);

            using var fileStream = File.Create(_currentFile);

            range.Save(fileStream, dataFormat);

            _isDirty = false;

            UpdateTitle();
            UpdateStatistics();
            UpdateStatus("Document salvat");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Documentul nu a putut fi salvat.\n\n{ex.Message}",
                "Eroare la salvare",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void Exit_Click(
        object sender,
        RoutedEventArgs e)
    {
        Close();
    }

    private void AcademicFormat_Click(
        object sender,
        RoutedEventArgs e)
    {
        ApplyAcademicDocumentFormat();

        _isDirty = true;

        UpdateTitle();
        UpdateStatus(
            "Format academic aplicat: Times New Roman 12, spațiere 1,5");
    }

    private void NormalText_Click(
        object sender,
        RoutedEventArgs e)
    {
        var paragraph = GetCurrentParagraph();

        if (paragraph is null)
            return;

        if (!IsParagraphEmpty(paragraph))
        {
            var newParagraph = CreateNormalParagraph();

            Editor.Document.Blocks.InsertAfter(
                paragraph,
                newParagraph);

            PlaceCaretAtEnd(newParagraph);
        }
        else
        {
            ApplyNormalParagraphStyle(paragraph);
            PlaceCaretAtEnd(paragraph);
        }

        _isDirty = true;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Text normal");
    }

    private void Chapter_Click(
        object sender,
        RoutedEventArgs e)
    {
        InsertHeading(1);
    }

    private void Subchapter_Click(
        object sender,
        RoutedEventArgs e)
    {
        InsertHeading(2);
    }

    private void InsertHeading(int level)
    {
        var currentParagraph = GetCurrentParagraph();

        if (currentParagraph is null)
            return;

        Paragraph headingParagraph;

        if (IsParagraphEmpty(currentParagraph))
        {
            headingParagraph = currentParagraph;
            headingParagraph.Inlines.Clear();
        }
        else
        {
            headingParagraph = new Paragraph();

            Editor.Document.Blocks.InsertAfter(
                currentParagraph,
                headingParagraph);
        }

        SetHeadingLevel(headingParagraph, level);

        ApplyHeadingStyle(
            headingParagraph,
            level);

        var prefix = GetNextHeadingPrefix(level);

        headingParagraph.Inlines.Add(
            new Run(prefix));

        PlaceCaretAtEnd(headingParagraph);

        _isDirty = true;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus(
            level == 1
                ? "Capitol nou"
                : "Subcapitol nou");
    }

    private void Renumber_Click(
        object sender,
        RoutedEventArgs e)
    {
        RenumberHeadings();

        _isDirty = true;

        UpdateTitle();
        UpdateStatistics();
        UpdateStatus("Numerotarea a fost recalculată");
    }

    private void Font_Changed(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (_suppressChanges)
            return;

        if (FontBox.SelectedItem is not FontFamily fontFamily)
            return;

        Editor.Selection.ApplyPropertyValue(
            TextElement.FontFamilyProperty,
            fontFamily);

        _isDirty = true;

        UpdateTitle();
        UpdateStatus("Font modificat");

        Editor.Focus();
    }

    private void Size_Changed(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (_suppressChanges)
            return;

        if (SizeBox.SelectedItem is not double fontSize)
            return;

        Editor.Selection.ApplyPropertyValue(
            TextElement.FontSizeProperty,
            fontSize);

        _isDirty = true;

        UpdateTitle();
        UpdateStatus("Dimensiune modificată");

        Editor.Focus();
    }

    private void ApplyAcademicDocumentFormat()
    {
        if (Editor?.Document is null)
            return;

        foreach (var paragraph in Editor.Document.Blocks.OfType<Paragraph>())
        {
            if (GetHeadingLevel(paragraph) == 0)
                ApplyNormalParagraphStyle(paragraph);
            else
                ApplyHeadingStyle(
                    paragraph,
                    GetHeadingLevel(paragraph));
        }
    }

    private Paragraph CreateNormalParagraph()
    {
        var paragraph = new Paragraph();

        ApplyNormalParagraphStyle(paragraph);

        SetHeadingLevel(paragraph, 0);

        return paragraph;
    }

    private void ApplyNormalParagraphStyle(
        Paragraph paragraph)
    {
        SetHeadingLevel(paragraph, 0);

        paragraph.FontFamily = new FontFamily(
            "Times New Roman");

        paragraph.FontSize = AcademicFontSize;
        paragraph.FontWeight = FontWeights.Normal;
        paragraph.FontStyle = FontStyles.Normal;

        paragraph.TextAlignment = TextAlignment.Justify;
        paragraph.LineHeight = AcademicLineHeight;
        paragraph.LineStackingStrategy =
            LineStackingStrategy.BlockLineHeight;

        paragraph.Margin = new Thickness(0);
        paragraph.Padding = new Thickness(0);
        paragraph.TextIndent = FirstLineIndent;
    }

    private void ApplyHeadingStyle(
        Paragraph paragraph,
        int level)
    {
        paragraph.FontFamily = new FontFamily(
            "Times New Roman");

        paragraph.FontWeight = FontWeights.Bold;
        paragraph.FontStyle = FontStyles.Normal;

        paragraph.TextAlignment =
            level == 1
                ? TextAlignment.Center
                : TextAlignment.Left;

        paragraph.LineHeight = AcademicLineHeight;
        paragraph.LineStackingStrategy =
            LineStackingStrategy.BlockLineHeight;

        paragraph.Margin = new Thickness(0);
        paragraph.Padding = new Thickness(0);
        paragraph.TextIndent = 0;

        paragraph.FontSize =
            level == 1
                ? 14
                : AcademicFontSize;
    }

    private Paragraph? GetCurrentParagraph()
    {
        if (Editor?.Selection?.Start is null)
            return null;

        var paragraph = Editor.Selection.Start.Paragraph;

        if (paragraph is not null)
            return paragraph;

        return Editor.Document.Blocks
            .OfType<Paragraph>()
            .LastOrDefault();
    }

    private static bool IsParagraphEmpty(
        Paragraph paragraph)
    {
        var range = new TextRange(
            paragraph.ContentStart,
            paragraph.ContentEnd);

        var text = range.Text ?? string.Empty;

        return string.IsNullOrWhiteSpace(text);
    }

    private void PlaceCaretAtEnd(
        Paragraph paragraph)
    {
        var caret = paragraph.ContentEnd;

        Editor.Selection.Select(caret, caret);
        Editor.Focus();
    }

    private string GetNextHeadingPrefix(
        int level)
    {
        var chapter = 0;
        var subchapter = 0;

        foreach (var paragraph in Editor.Document.Blocks
                     .OfType<Paragraph>())
        {
            var headingLevel = GetHeadingLevel(paragraph);

            if (headingLevel == 1)
            {
                chapter++;
                subchapter = 0;
            }
            else if (headingLevel == 2)
            {
                subchapter++;
            }
        }

        if (level == 1)
            return $"Capitolul {chapter + 1}. ";

        if (chapter == 0)
            chapter = 1;

        return $"{chapter}.{subchapter + 1}. ";
    }

    private void RenumberHeadings()
    {
        var chapter = 0;
        var subchapter = 0;

        foreach (var paragraph in Editor.Document.Blocks
                     .OfType<Paragraph>())
        {
            var level = GetHeadingLevel(paragraph);

            if (level == 0)
                continue;

            if (level == 1)
            {
                chapter++;
                subchapter = 0;
            }
            else if (level == 2)
            {
                if (chapter == 0)
                    chapter = 1;

                subchapter++;
            }

            var prefix = level == 1
                ? $"Capitolul {chapter}. "
                : $"{chapter}.{subchapter}. ";

            ReplaceHeadingPrefix(
                paragraph,
                prefix);

            ApplyHeadingStyle(
                paragraph,
                level);
        }
    }

    private void ReplaceHeadingPrefix(
        Paragraph paragraph,
        string prefix)
    {
        if (paragraph.Inlines.FirstInline is Run firstRun)
        {
            firstRun.Text = prefix;
            return;
        }

        paragraph.Inlines.InsertBefore(
            paragraph.Inlines.FirstInline,
            new Run(prefix));
    }

    private void RestoreHeadingMetadata()
    {
        foreach (var paragraph in Editor.Document.Blocks
                     .OfType<Paragraph>())
        {
            var textRange = new TextRange(
                paragraph.ContentStart,
                paragraph.ContentEnd);

            var text = textRange.Text ?? string.Empty;

            if (Regex.IsMatch(
                    text,
                    @"^\s*Capitolul\s+\d+\.\s"))
            {
                SetHeadingLevel(paragraph, 1);
                ApplyHeadingStyle(paragraph, 1);
            }
            else if (Regex.IsMatch(
                         text,
                         @"^\s*\d+\.\d+\.\s"))
            {
                SetHeadingLevel(paragraph, 2);
                ApplyHeadingStyle(paragraph, 2);
            }
            else
            {
                SetHeadingLevel(paragraph, 0);
                ApplyNormalParagraphStyle(paragraph);
            }
        }
    }

    private void LightTheme_Click(
        object sender,
        RoutedEventArgs e)
    {
        ApplyLightTheme();
    }

    private void DarkTheme_Click(
        object sender,
        RoutedEventArgs e)
    {
        ApplyDarkTheme();
    }

    private void ApplyLightTheme()
    {
        _isDarkTheme = false;

        MainRoot.Background = new SolidColorBrush(
            Color.FromRgb(243, 244, 246));

        MainMenu.Background = Brushes.White;
        ToolbarBorder.Background = Brushes.White;
        AcademicToolbar.Background = new SolidColorBrush(
            Color.FromRgb(232, 238, 249));

        Sidebar.Background = Brushes.White;
        EditorBorder.Background = Brushes.White;
        Editor.Background = Brushes.White;
        StatusBorder.Background = Brushes.White;

        Editor.Foreground = new SolidColorBrush(
            Color.FromRgb(31, 41, 55));

        UpdateStatus("Tema deschisă activată");
    }

    private void ApplyDarkTheme()
    {
        _isDarkTheme = true;

        var darkBackground = new SolidColorBrush(
            Color.FromRgb(15, 23, 42));

        var darkPanel = new SolidColorBrush(
            Color.FromRgb(30, 41, 59));

        var editorBackground = new SolidColorBrush(
            Color.FromRgb(248, 250, 252));

        MainRoot.Background = darkBackground;

        MainMenu.Background = darkPanel;
        ToolbarBorder.Background = darkPanel;
        AcademicToolbar.Background = new SolidColorBrush(
            Color.FromRgb(51, 65, 85));

        Sidebar.Background = darkPanel;
        EditorBorder.Background = darkPanel;
        Editor.Background = editorBackground;
        StatusBorder.Background = darkPanel;

        Editor.Foreground = new SolidColorBrush(
            Color.FromRgb(15, 23, 42));

        UpdateStatus("Tema întunecată activată");
    }

    private bool ConfirmDiscardChanges()
    {
        if (!_isDirty)
            return true;

        var result = MessageBox.Show(
            "Documentul conține modificări nesalvate. Doriți să le salvați?",
            "Modificări nesalvate",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Warning);

        if (result == MessageBoxResult.Cancel)
            return false;

        if (result == MessageBoxResult.No)
            return true;

        if (result == MessageBoxResult.Yes)
        {
            Save_Click(
                this,
                new RoutedEventArgs());

            return !_isDirty;
        }

        return false;
    }

    private void Window_Closing(
        object? sender,
        CancelEventArgs e)
    {
        if (!ConfirmDiscardChanges())
            e.Cancel = true;
    }

    private void About_Click(
        object sender,
        RoutedEventArgs e)
    {
        MessageBox.Show(
            "MyWord\n\n"
            + "Editor academic pentru studenți\n"
            + "Versiunea 1.0",
            "Despre MyWord",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void UpdateTitle()
    {
        var fileName = string.IsNullOrWhiteSpace(_currentFile)
            ? "Document nou"
            : Path.GetFileName(_currentFile);

        var modifiedMark = _isDirty
            ? " *"
            : string.Empty;

        Title = $"MyWord - {fileName}{modifiedMark}";
    }

    private void UpdateStatistics()
    {
        if (Editor?.Document is null)
            return;

        var range = new TextRange(
            Editor.Document.ContentStart,
            Editor.Document.ContentEnd);

        var text = range.Text ?? string.Empty;

        var words = text
            .Split(
                new[] { ' ', '\r', '\n', '\t' },
                StringSplitOptions.RemoveEmptyEntries)
            .Length;

        var characters = text
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Length;

        if (WordsText is not null)
            WordsText.Text = $"Cuvinte: {words}";

        if (CharactersText is not null)
            CharactersText.Text = $"Caractere: {characters}";

        if (FileText is not null)
        {
            FileText.Text = string.IsNullOrWhiteSpace(_currentFile)
                ? "Document nou"
                : Path.GetFileName(_currentFile);
        }
    }

    private void UpdateStatus(string message)
    {
        if (StatusText is not null)
            StatusText.Text = message;
    }
}