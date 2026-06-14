using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SlideAudienceAddIn.Models;

namespace SlideAudienceAddIn.Settings
{
    public partial class SettingsWindow : Window
    {
        private readonly AppSettings _settings;

        public SettingsWindow(AppSettings settings)
        {
            _settings = settings;
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            EnableAddInCheckBox.IsChecked = _settings.Enabled;
            EnableApiCheckBox.IsChecked = _settings.Gemini.EnableApi;
            ModelTextBox.Text = _settings.Gemini.Model;
            TimeoutTextBox.Text = _settings.Gemini.TimeoutSeconds.ToString();
            FontSizeTextBox.Text = _settings.Overlay.FontSize.ToString("0");
            EnableLoggingCheckBox.IsChecked = _settings.Experiment.EnableLogging;
            SaveSlideTextCheckBox.IsChecked = _settings.Experiment.SaveSlideTextToLog;

            SelectComboBoxItem(ModeComboBox, _settings.Comments.Mode);
            SelectComboBoxItem(MaxCommentsComboBox, _settings.Comments.MaxCommentsPerSlide.ToString());
            SelectComboBoxItem(PositionComboBox, _settings.Overlay.Position);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(TimeoutTextBox.Text, out var timeoutSeconds) || timeoutSeconds < 5)
            {
                MessageBox.Show("Timeout seconds must be 5 or greater.", "SlideAudience Settings");
                return;
            }

            if (!double.TryParse(FontSizeTextBox.Text, out var fontSize) || fontSize < 12 || fontSize > 48)
            {
                MessageBox.Show("Font size must be between 12 and 48.", "SlideAudience Settings");
                return;
            }

            _settings.Enabled = EnableAddInCheckBox.IsChecked == true;
            _settings.Gemini.EnableApi = EnableApiCheckBox.IsChecked == true;
            _settings.Gemini.Model = ModelTextBox.Text.Trim();
            _settings.Gemini.TimeoutSeconds = timeoutSeconds;
            _settings.Comments.Mode = SelectedText(ModeComboBox, _settings.Comments.Mode);
            _settings.Comments.MaxCommentsPerSlide = int.Parse(SelectedText(MaxCommentsComboBox, "3"));
            _settings.Overlay.FontSize = fontSize;
            _settings.Overlay.Position = SelectedText(PositionComboBox, "BottomRight");
            _settings.Experiment.EnableLogging = EnableLoggingCheckBox.IsChecked == true;
            _settings.Experiment.SaveSlideTextToLog = SaveSlideTextCheckBox.IsChecked == true;
            _settings.SaveLocal();

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private static void SelectComboBoxItem(ComboBox comboBox, string value)
        {
            foreach (var item in comboBox.Items.OfType<ComboBoxItem>())
            {
                if (string.Equals(Convert.ToString(item.Content), value, StringComparison.OrdinalIgnoreCase))
                {
                    comboBox.SelectedItem = item;
                    return;
                }
            }

            comboBox.SelectedIndex = 0;
        }

        private static string SelectedText(ComboBox comboBox, string fallback)
        {
            return comboBox.SelectedItem is ComboBoxItem item
                ? Convert.ToString(item.Content)
                : fallback;
        }
    }
}
