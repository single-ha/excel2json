using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

/// <summary>
/// 通过文件夹选择文件 直接选择.obj文件 并记录路径
/// </summary>
public class FileChoose_Dll : IFileChoose
{
    private enum OpenFileNameFlags
    {
        /// <summary>
        /// 隐藏只读复选框
        /// </summary>
        OFN_HIDEREADONLY = 0x4,
        /// <summary>
        /// 显示隐藏文件
        /// </summary>
        OFN_FORCESHOWHIDDEN = 0x10000000,
        /// <summary>
        /// 可多选
        /// </summary>
        OFN_ALLOWMULTISELECT = 0x200,
        /// <summary>
        /// 资源管理器
        /// </summary>
        OFN_EXPLORER = 0x80000,
        OFN_FILEMUSTEXIST = 0x1000,
        OFN_PATHMUSTEXIST = 0x800,
        OFN_NOCHANGEDIR=0x8,
    }
    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    public string[] OpenFileDialog(string dialogTitle, string startPath, string filter, bool showHidden, bool allowMultiSelect)
    {
        const int MAX_FILE_LENGTH = 2048;
        OpenFileName ofn = new OpenFileName();
        ofn.structSize = Marshal.SizeOf(ofn);
        ofn.filter = filter.Replace("|", "\0") + "\0";
        ofn.fileTitle = new String(new char[MAX_FILE_LENGTH]);
        ofn.maxFileTitle = ofn.fileTitle.Length;
        ofn.initialDir = startPath;
        ofn.title = dialogTitle;
        ofn.flags = (int)OpenFileNameFlags.OFN_EXPLORER | (int)OpenFileNameFlags.OFN_FILEMUSTEXIST | (int)OpenFileNameFlags.OFN_PATHMUSTEXIST| (int)OpenFileNameFlags.OFN_NOCHANGEDIR;
        // ofn.dlgOwner = GetActiveWindow();
        // Create buffer for file names
        string fileNames = new String(new char[MAX_FILE_LENGTH]);
        ofn.file =  Marshal.StringToBSTR(fileNames);
        ofn.maxFile = fileNames.Length;

        if (showHidden)
        {
            ofn.flags |= (int)OpenFileNameFlags.OFN_FORCESHOWHIDDEN;
        }

        if (allowMultiSelect)
        {
            ofn.flags |= (int)OpenFileNameFlags.OFN_ALLOWMULTISELECT;
        }

        if (DllOpenFileDialog.GetOpenFileName(ofn))
        {
            List<string> selectedFilesList = new List<string>();
            
            long pointer = (long)ofn.file;
            string file = Marshal.PtrToStringAuto(ofn.file);
            // Retrieve file names
            while (file.Length > 0)
            {
                selectedFilesList.Add(file);
            
                pointer += file.Length * Marshal.SystemDefaultCharSize + Marshal.SystemDefaultCharSize;
                ofn.file = (IntPtr)pointer;
                file = Marshal.PtrToStringAuto(ofn.file);
            }
            
            if (selectedFilesList.Count == 1)
            {
                // Only one file selected with full path
                return selectedFilesList.ToArray();
            }
            else
            {
                // Multiple files selected, add directory
                string[] selectedFiles = new string[selectedFilesList.Count - 1];
            
                for (int i = 0; i < selectedFiles.Length; i++)
                {
                    selectedFiles[i] = selectedFilesList[0] + "\\" + selectedFilesList[i + 1];
                }
            
                return selectedFiles;
            }
        }
        else
        {
            // "Cancel" pressed
            return null;
        }
    }

    public string OpenFolderDialog(string description)
    {
        OpenDialogDir openDir = new OpenDialogDir();
        openDir.pszDisplayName = new string(new char[2000]);
        openDir.lpszTitle = description;
        openDir.ulFlags = 1; // BIF_NEWDIALOGSTYLE | BIF_EDITBOX;
        IntPtr pidl = DllOpenFileDialog.SHBrowseForFolder(openDir);

        char[] path = new char[2000];
        for (int i = 0; i < 2000; i++)
            path[i] = '\0';
        if (DllOpenFileDialog.SHGetPathFromIDList(pidl, path))
        {
            string str = new string(path);
            string DirPath = str.Substring(0, str.IndexOf('\0'));
            return DirPath;
        }

        return "";
    }

    public void OpenFolder(string folderPath)
    {
        System.Diagnostics.Process.Start("explorer.exe", folderPath);
    }
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class OpenFileName
{
    public int structSize = 0;
    public IntPtr dlgOwner = IntPtr.Zero;
    public IntPtr instance = IntPtr.Zero;
    public String filter = null;
    public String customFilter = null;
    public int maxCustFilter = 0;
    public int filterIndex = 0;
    public IntPtr file = IntPtr.Zero;
    public int maxFile = 0;
    public String fileTitle = null;
    public int maxFileTitle = 0;
    public String initialDir = null;
    public String title = null;
    public int flags = 0;
    public short fileOffset = 0;
    public short fileExtension = 0;
    public String defExt = null;
    public IntPtr custData = IntPtr.Zero;
    public IntPtr hook = IntPtr.Zero;
    public String templateName = null;
    public IntPtr reservedPtr = IntPtr.Zero;
    public int reservedInt = 0;
    public int flagsEx = 0;
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class OpenDialogDir
{
    public IntPtr hwndOwner = IntPtr.Zero;
    public IntPtr pidlRoot = IntPtr.Zero;
    public String pszDisplayName = "123";
    public String lpszTitle = null;
    public UInt32 ulFlags = 0;
    public IntPtr lpfn = IntPtr.Zero;
    public IntPtr lParam = IntPtr.Zero;
    public int iImage = 0;
}

public class DllOpenFileDialog
{
    [DllImport("Comdlg32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool GetOpenFileName([In, Out] OpenFileName ofn);

    [DllImport("Comdlg32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool GetSaveFileName([In, Out] OpenFileName ofn);

    [DllImport("shell32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SHBrowseForFolder([In, Out] OpenDialogDir ofn);

    [DllImport("shell32.dll", SetLastError = true, ThrowOnUnmappableChar = true, CharSet = CharSet.Auto)]
    public static extern bool SHGetPathFromIDList([In] IntPtr pidl, [In, Out] char[] fileName);
}