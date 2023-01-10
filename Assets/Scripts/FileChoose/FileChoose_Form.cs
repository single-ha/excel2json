using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

public class FileChoose_Form: IFileChoose
{
    private System.Threading.Thread CloseOops;
    public bool ThreadRunning = true;
    public const int WM_SYSCOMMAND = 0x0112;
    public const int SC_CLOSE = 0xF060;

    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

    public void OopsWindowsThreadStart()
    {
        CloseOops = new System.Threading.Thread(ClearOopsWindows) { IsBackground = true };
        ThreadRunning = true;
        CloseOops.Start();
    }
    private void ClearOopsWindows()
    {
        while (ThreadRunning)
        {
            FindAndCloseWindow();
        }
    }
    public static void FindAndCloseWindow()
    {
        IntPtr lHwnd = FindWindow(null, "Oops");
        if (lHwnd != IntPtr.Zero)
        {
            SendMessage(lHwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
        }
    }

    public string[] OpenFileDialog(string dialogTitle, string startPath, string filter, bool showHidden, bool allowMultiSelect)
    {
        OopsWindowsThreadStart();
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Multiselect = true;
        openFileDialog.Title = "选择文件";
        openFileDialog.Filter = "选择文件(.xlsx,.xls)|*.xls";
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            return openFileDialog.FileNames;
        }

        ThreadRunning = false;
        return null;
    }

    public string OpenFolderDialog(string description)
    {
        OopsWindowsThreadStart();
        FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
        folderBrowserDialog.Description = "选择文件夹";
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            return folderBrowserDialog.SelectedPath;
        }
        ThreadRunning = false;
        return null;
    }

    public void OpenFolder(string folderPath)
    {
        System.Diagnostics.Process.Start("explorer.exe", folderPath);
    }
}