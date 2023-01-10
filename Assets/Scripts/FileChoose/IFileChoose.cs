public interface IFileChoose
{
    string[] OpenFileDialog(string dialogTitle, string startPath, string filter, bool showHidden, bool allowMultiSelect);
    string OpenFolderDialog(string description);
    void OpenFolder(string folderPath);
}