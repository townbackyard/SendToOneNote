using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SendToOneNote.Core.Picker;

namespace SendToOneNote;

public partial class PickerWindow : Window
{
    private readonly SectionPickerViewModel _vm;
    public PickerItem? Selected { get; private set; }

    public PickerWindow(SectionPickerViewModel vm, string emailSubject)
    {
        InitializeComponent();
        _vm = vm;
        Title = $"Send to OneNote — {emailSubject}";
        Results.ItemsSource = _vm.Filter("");
        if (Results.Items.Count > 0) Results.SelectedIndex = 0;
        Loaded += (_, _) => SearchBox.Focus();
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        Results.ItemsSource = _vm.Filter(SearchBox.Text);
        if (Results.Items.Count > 0) Results.SelectedIndex = 0;
    }

    private void SearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Down && Results.Items.Count > 0)
        {
            Results.SelectedIndex = Math.Min(Results.SelectedIndex + 1, Results.Items.Count - 1);
            e.Handled = true;
        }
        else if (e.Key == Key.Up && Results.SelectedIndex > 0)
        {
            Results.SelectedIndex--;
            e.Handled = true;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        Selected = Results.SelectedItem as PickerItem;
        if (Selected is null) return;
        DialogResult = true;
    }
}
