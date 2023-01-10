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
    public Button Btn_SelectFolder;

    public Button Btn_SelectFiles;

    public ScrollRect SelectFiles;

    public Text Tex_SelectFilesPath;

    public Button Btn_OutPutFolder;
    public Button Btn_OpenOutPut;
    public Button Btn_Clear;
    public Text Tex_OutPutFolderPath;

    public Button Btn_Create;
    public Button Btn_ShowLog;
    public Transform messageRoot;
    public Text showMessage;
    public Button Btn_CloseMessage;
    public Button Btn_clearLog;
    public Transform processRoot;
    public Image process;

    private IFileChoose fileChoose;
    StringBuilder sb_Select = new StringBuilder();
    private string folderPath;
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
        if (Btn_SelectFolder != null)
        {
            Btn_SelectFolder.onClick.AddListener(OnClickSelectFolder);
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
        folderPath = null;
        filesPath = null;
        RefreshPahtsShow();
    }

    private void InitData()
    {
        sw = new Stopwatch();
        fileChoose = new FileChoose_Dll();
        log_List = new List<string>();
        allfiles = new HashSet<string>();
        folderPath = PlayerPrefs.GetString("folderPath", "");
        outPutPath = PlayerPrefs.GetString("outPutPath", "");
        Excel2Json.Init();
    }

    private void OnClickShowLog()
    {
        StringBuilder sb = new StringBuilder();
        for (int i = log_List.Count-1; i >= 0; i--)
        {
            sb.AppendLine(log_List[i]);
        }
        Instance.ShowLogPanel(sb.ToString());
    }

    private void OnClickCreate()
    {
        if (string.IsNullOrEmpty(outPutPath))
        {
            Log("δѡ�����·��");
        }

        allfiles.Clear();
        CollectFiles();
        if (allfiles.Count <= 0)
        {
            LogError("��ѡ��Excel�ļ�");
            ShowLogPanel("��ѡ��Excel�ļ�");
            return;
        }
        StartCoroutine(StartExport()) ;
    }

    private IEnumerator StartExport()
    {
        count = 0;
        processValue = 0;
        sw.Restart();
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
                StartCoroutine(Excel2Json.Excel2Json_File(fileInfo, outPutPath, delegate ()
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

    private void CollectFiles()
    {
        if (!string.IsNullOrEmpty(folderPath))
        {
            GetExcelFileInFolder();
        }

        if (filesPath != null)
        {
            for (int i = 0; i < filesPath.Length; i++)
            {
                allfiles.Add(filesPath[i]);
            }
        }
    }

    private void OnClickOutPut()
    {
        outPutPath = fileChoose.OpenFolderDialog("ѡ�����·��");
        PlayerPrefs.SetString("outPutPath", outPutPath);
        RefreshOutputShow();
    }

    private void RefreshOutputShow()
    {
        Tex_OutPutFolderPath.text = outPutPath;
    }
    private void OnClickSelectFiles()
    {
        string Title = "ѡ���ļ�";
        string startPath = Application.streamingAssetsPath.Replace('/', '\\'); //Ĭ��·��
        string filter = "Excel�ļ�(.xls,.xlsx)|*.xlsx;*.xls";
        filesPath = fileChoose.OpenFileDialog(Title, startPath, filter, false, true);
        RefreshPahtsShow();
    }

    private void OnClickSelectFolder()
    {
        folderPath = fileChoose.OpenFolderDialog("ѡ���ļ���");
        RefreshPahtsShow();
        PlayerPrefs.SetString("folderPath",folderPath);
    }

    private void RefreshPahtsShow()
    {
        sb_Select.Clear();
        if (!string.IsNullOrEmpty(folderPath))
        {
            sb_Select.AppendLine($"<color=#495BEC>{folderPath}</color>");
        }

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
        if (!Directory.Exists(folderPath))
        {
            return;
        }

        DirectoryInfo di = new DirectoryInfo(folderPath);
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
        if (dir.FullName==folderPath)
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
                Log($"��ʱ:{sw.ElapsedMilliseconds/(float)1000}s");
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
    //     //ѡ��ĳһ�ļ�
    //     if (GUI.Button(new Rect(10, 10, 100, 50), "ChooseFile"))
    //     {
    //         string Title = "ѡ���ļ�";
    //         string startPath = Application.streamingAssetsPath.Replace('/', '\\'); //Ĭ��·��
    //         string filter = "Excel�ļ�(.xls,.xlsx)|*.xlsx;*.xls";
    //         var paths = fileChoose.OpenFileDialog(Title, startPath, filter, false, true);
    //         if (paths != null)
    //         {
    //             for (int i = 0; i < paths.Length; i++)
    //             {
    //                 Debug.Log(paths[i]);
    //             }
    //         }
    //     }
    //     //ѡ��ĳһ�ļ���
    //     if (GUI.Button(new Rect(10, 100, 100, 50), "ChooseDirectory"))
    //     {
    //         var path = fileChoose.OpenFolderDialog("ѡ���ļ���");
    //         Debug.Log(path);
    //     }
    // }
}