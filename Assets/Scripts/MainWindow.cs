using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using UnityEngine;
using UnityEngine.UI;
using Button = UnityEngine.UI.Button;
using Debug = UnityEngine.Debug;

public class MainWindow : MonoBehaviour
{
    public Button Btn_Clear;
    public Button Btn_ShowLog;
    public Button Btn_SelectFiles;
    public Button Btn_Create;

    public Text Tex_InputRoot;
    public Button Btn_SelectInPutRoot;
    public Button Btn_OpenInput;


    public ScrollRect SelectFiles;
    public Text Tex_SelectFilesPath;
    public Button Btn_OutPutFolder;
    public Button Btn_OpenOutPut;
    public Text Tex_OutPutFolderPath;

    public Transform messageRoot;
    public Text showMessage;
    public Button Btn_CloseMessage;
    public Button Btn_clearLog;
    public Transform processRoot;
    public Image process;

    private IFileChoose fileChoose;
    StringBuilder sb_Select = new StringBuilder();
    private string inputRoot;
    private string[] filesPath;
    private string outPutPath;
    private HashSet<string> allfiles;
    private static List<string> log_List;
    private int count;
    private float processValue;
    
    public static MainWindow Instance;

    private Stopwatch sw;
    // Start is called before the first frame update
    void Start()
    {
        Instance = this;
        if (Btn_SelectInPutRoot != null)
        {
            Btn_SelectInPutRoot.onClick.AddListener(OnClickSelectFolder);
        }

        if (Btn_OpenInput!=null)
        {
            Btn_OpenInput.onClick.AddListener(delegate()
            {
                if (!string.IsNullOrEmpty(inputRoot))
                {
                    fileChoose.OpenFolder(inputRoot);
                }
            });
        }
        if (Btn_SelectFiles != null)
        {
            Btn_SelectFiles.onClick.AddListener(OnClickSelectFiles);
        }

        if (Btn_OutPutFolder != null)
        {
            Btn_OutPutFolder.onClick.AddListener(OnClickOutPut);
        }

        if (Btn_Create != null)
        {
            Btn_Create.onClick.AddListener(OnClickCreate);
        }

        if (Btn_ShowLog != null)
        {
            Btn_ShowLog.onClick.AddListener(OnClickShowLog);
        }

        if (Btn_Clear != null)
        {
            Btn_Clear.onClick.AddListener(OnClear);
        }
        if (Btn_CloseMessage != null)
        {
            Btn_CloseMessage.onClick.AddListener(delegate() { messageRoot?.gameObject.SetActive(false); });
        }

        if (Btn_clearLog!=null)
        {
            Btn_clearLog.onClick.AddListener(delegate()
            {
                log_List.Clear();
                OnClickShowLog();
            });
        }

        if (Btn_OpenOutPut!=null)
        {
            Btn_OpenOutPut.onClick.AddListener(delegate()
            {
                if (!string.IsNullOrEmpty(outPutPath))
                {
                    fileChoose.OpenFolder(outPutPath);
                }
            });
        }
        messageRoot?.gameObject.SetActive(false);
        processRoot?.gameObject.SetActive(false);
        InitData();
        RefreshPahtsShow();
        RefreshOutputShow();
    }

    private void OnClear()
    {
        inputRoot = null;
        filesPath = null;
        RefreshPahtsShow();
    }

    private void InitData()
    {
        sw = new Stopwatch();
        fileChoose = new FileChoose_Dll();
        log_List = new List<string>();
        allfiles = new HashSet<string>();
        inputRoot = PlayerPrefs.GetString("inputRoot", "");
        outPutPath = PlayerPrefs.GetString("outPutPath", "");
        Excel2Json.Init();
    }

    private void OnClickShowLog()
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < log_List.Count; i++)
        {
            sb.AppendLine(log_List[i]);
        }
        // for (int i = log_List.Count-1; i >= 0; i--)
        // {
        //     sb.AppendLine(log_List[i]);
        // }
        Instance.ShowLogPanel(sb.ToString());
    }

    private void OnClickCreate()
    {
        if (string.IsNullOrEmpty(outPutPath))
        {
            Log("未选择输出路径");
        }
        allfiles.Clear();
        CollectFiles();
        if (allfiles.Count <= 0)
        {
            LogError("请选择Excel文件");
            ShowLogPanel("请选择Excel文件");
            return;
        }
        StartCoroutine(StartExport()) ;
    }

    private IEnumerator StartExport()
    {
        count = 0;
        processValue = 0;
        sw.Restart();
        if (filesPath==null||filesPath.Length<0)
        {
            //全量模式
            ClearFolder(outPutPath);
        }
        var arrs = allfiles.ToArray();
        StartCoroutine(ShowProcess());
        if (arrs.Length>0)
        {
            for (int i = 0; i < arrs.Length; i++)
            {
                var excelPath = arrs[i];
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!CheckFile(fileInfo))
                {
                    continue;
                }

                string output = outPutPath;
                if (fileInfo.DirectoryName!=null && fileInfo.DirectoryName!=inputRoot)
                {
                    output = fileInfo.DirectoryName.Replace(inputRoot, outPutPath);
                }
                else
                {
                    
                }
                StartCoroutine(Excel2Json.Excel2Json_File(fileInfo, output, delegate ()
                {
                    count++;
                    processValue = count / (float)allfiles.Count;
                }));
                yield return null;
            }
        }
        else
        {
            processValue = 1;
        }

        // StartCoroutine(Excel2Json.Excel2Json_Files(allfiles.ToArray(), outPutPath, delegate()
        // {
        //     count++;
        //     processValue = count / (float)allfiles.Count;
        // }));
        // Excel2Json.Excel2Json_Files(allfiles.ToArray(), outPutPath, delegate()
        // {
        //     count++;
        //     processValue = count / (float)allfiles.Count;
        // });
    }

    private void ClearFolder(string path)
    {
        if (!Directory.Exists(path))
        {
            return;
        }
        var files = Directory.GetFiles(path,"*.*",SearchOption.AllDirectories);
        for (int i = 0; i < files.Length; i++)
        {
            File.Delete(files[i]);
        }

        var dirs = Directory.GetDirectories(path,"*",SearchOption.AllDirectories);
        for (int i = 0; i < dirs.Length; i++)
        {
            Directory.Delete(dirs[i]);
        }
    }
    private void CollectFiles()
    {
        if (filesPath != null&&filesPath.Length>0)
        {
            for (int i = 0; i < filesPath.Length; i++)
            {
                allfiles.Add(filesPath[i]);
            }
            return;//当有单个表格选中的时候为增量转化否则转化根目录下的所有表格
        }
        if (!string.IsNullOrEmpty(inputRoot))
        {
            GetExcelFileInFolder();
        }
    }

    private void OnClickOutPut()
    {
        outPutPath = fileChoose.OpenFolderDialog("选择输出路径");
        PlayerPrefs.SetString("outPutPath", outPutPath);
        RefreshOutputShow();
    }

    private void RefreshOutputShow()
    {
        Tex_OutPutFolderPath.text = outPutPath;
    }
    private void OnClickSelectFiles()
    {
        string Title = "选择文件";
        string startPath = Application.streamingAssetsPath.Replace('/', '\\'); //默认路径
        string filter = "Excel文件(.xls,.xlsx)|*.xlsx;*.xls";
        filesPath = fileChoose.OpenFileDialog(Title, startPath, filter, false, true);
        RefreshPahtsShow();
    }

    private void OnClickSelectFolder()
    {
        inputRoot = fileChoose.OpenFolderDialog("选择文件夹");
        RefreshPahtsShow();
        PlayerPrefs.SetString("inputRoot",inputRoot);
    }

    private void RefreshPahtsShow()
    {
        if (!string.IsNullOrEmpty(inputRoot))
        {
            Tex_InputRoot.text = inputRoot;
            // sb_Select.AppendLine($"<color=#495BEC>{inputRoot}</color>");
        }

        sb_Select.Clear();
        if (filesPath != null)
        {
            for (int i = 0; i < filesPath.Length; i++)
            {
                sb_Select.AppendLine(filesPath[i]);
            }
        }

        Tex_SelectFilesPath.text = sb_Select.ToString();
        // LayoutRebuilder.ForceRebuildLayoutImmediate(Tex_SelectFilesPath.GetComponent<RectTransform>());
        // SelectFiles.verticalNormalizedPosition = 0;
    }

    private void GetExcelFileInFolder()
    {
        if (!Directory.Exists(inputRoot))
        {
            return;
        }

        DirectoryInfo di = new DirectoryInfo(inputRoot);
        var files = di.GetFiles("*.xls", SearchOption.AllDirectories);
        for (int i = 0; i < files.Length; i++)
        {
            if (!CheckFile(files[i]))
            {
                continue;
            }
            if (!CheckDirectory(files[i].Directory))
            {
                continue;
            }
            allfiles.Add(files[i].FullName);
        }
    }

    private bool CheckFile(FileInfo file)
    {
        if (file.Name.StartsWith("~"))
        {
            return false;
        }

        if (file.Name.StartsWith("_"))
        {
            return false;
        }
        return true;
    }
    private bool CheckDirectory(DirectoryInfo dir)
    {
        if (dir.FullName==inputRoot)
        {
            return true;
        }
        else
        {
            if (dir.Name.StartsWith("_"))
            {
                return false;
            }
            else
            {
                return CheckDirectory(dir.Parent);
            }
        }
    }
    public void ShowLogPanel(string message)
    {
        if (messageRoot != null)
        {
            messageRoot.gameObject.SetActive(true);
            showMessage.text = message;
        }
    }

    private IEnumerator ShowProcess()
    {
        processRoot.gameObject.SetActive(true);
        while (processRoot != null)
        {
            RefreshProcess();
            if (processValue>=1)
            {
                sw.Stop();
                Log($"用时:{sw.ElapsedMilliseconds/(float)1000}s");
                CloseProcess();
                ShowLogPanel("complat!!!");
                yield break;
            }
            yield return null;
        }
    }

    private void RefreshProcess()
    {
        if (processRoot != null)
        {
            var scale = process.transform.localScale;
            scale.x = processValue;
            process.transform.localScale = scale;
        }
    }
    private void CloseProcess()
    {
        processRoot.gameObject.SetActive(false);
    }

    public static void LogError(string log)
    {
        Debug.LogError(log);
        log_List.Add($"<color=#B70000>{log}</color>");
    }

    public static void Log(string log)
    {
        Debug.Log(log);
        log_List.Add($"{log}");
    }
    // void OnGUI()
    // {
    //     //选择某一文件
    //     if (GUI.Button(new Rect(10, 10, 100, 50), "ChooseFile"))
    //     {
    //         string Title = "选择文件";
    //         string startPath = Application.streamingAssetsPath.Replace('/', '\\'); //默认路径
    //         string filter = "Excel文件(.xls,.xlsx)|*.xlsx;*.xls";
    //         var paths = fileChoose.OpenFileDialog(Title, startPath, filter, false, true);
    //         if (paths != null)
    //         {
    //             for (int i = 0; i < paths.Length; i++)
    //             {
    //                 Debug.Log(paths[i]);
    //             }
    //         }
    //     }
    //     //选择某一文件夹
    //     if (GUI.Button(new Rect(10, 100, 100, 50), "ChooseDirectory"))
    //     {
    //         var path = fileChoose.OpenFolderDialog("选择文件夹");
    //         Debug.Log(path);
    //     }
    // }
}