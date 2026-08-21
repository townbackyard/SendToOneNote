using System.IO;
using System.Windows;
using Microsoft.Win32;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Storage;

namespace SendToOneNote;

public partial class FirstRunWindow : Window
{
    private readonly ITokenProvider _tokens;
    private readonly AppSettings _settings;
    public bool Completed { get; private set; }

    public FirstRunWindow(ITokenProvider tokens, AppSettings settings)
    {
        InitializeComponent();
        _tokens = tokens;
        _settings = settings;
        FolderBox.Text = settings.DropFolder ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "SendToOneNote Drop");
    }

    private async void SignIn_Click(object sender, RoutedEventArgs e)
    {
        SignInBtn.IsEnabled = false;
        try
        {
            await _tokens.GetAccessTokenAsync(interactiveAllowed: true);
            SignedInAs.Text = $"Signed in as {_tokens.SignedInUser}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Sign-in failed");
        }
        finally { SignInBtn.IsEnabled = true; }
    }

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog();
        if (dlg.ShowDialog(this) == true) FolderBox.Text = dlg.FolderName;
    }

    private void Finish_Click(object sender, RoutedEventArgs e)
    {
        if (_tokens.SignedInUser is null)
        {
            MessageBox.Show(this, "Please sign in first.", "SendToOneNote");
            return;
        }
        if (string.IsNullOrWhiteSpace(FolderBox.Text))
        {
            MessageBox.Show(this, "Please choose a drop folder.", "SendToOneNote");
            return;
        }
        _settings.DropFolder = FolderBox.Text;
        Directory.CreateDirectory(_settings.DropFolder);
        if (StartupBox.IsChecked == true)
        {
            try { CreateStartupShortcut(); }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Couldn't create the Startup shortcut: {ex.Message}\nYou can add one manually later.",
                    "SendToOneNote");
            }
        }
        Completed = true;
        Close();
    }

    private static void CreateStartupShortcut()
    {
        var startup = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        var lnk = Path.Combine(startup, "SendToOneNote.lnk");
        var exe = Environment.ProcessPath!;
        dynamic? shell = null;
        dynamic? sc = null;
        try
        {
            shell = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell")!)!;
            sc = shell.CreateShortcut(lnk);
            sc.TargetPath = exe;
            sc.WorkingDirectory = Path.GetDirectoryName(exe);
            sc.Save();
        }
        finally
        {
            if (sc is not null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sc);
            if (shell is not null) System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
        }
    }
}
