using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using LitJson;
using OfficeOpenXml;

/// <summary>
/// 读取excel表格
/// </summary>
/*
 * |----|------|----|----|
 * |id  |name  | num|    |变量名称
 * |----|------|----|----|
 * | int|string|int |    |数值类型(int,string,double,bool,Array)
 * |----|------|----|----|
 * |1   |张三   |101 |    |
 * |----|------|----|----|
 * 注意:如果excel里有Table则转换成字典,如果没有则转换成数组.比如
 * 字典:
 * {
 *      "1":{
 *              "name":"张三".
 *              "num":101
 *          },
 * }
 *没有table,数组:
 * [{
 *      "id":1,
 *      "name":"张三",
 *      "num":101
 * }]
 */
public static class Excel2Json
{
    public static void Init()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    public static IEnumerator Excel2Json_File(string excelPath,string outPutFloder=null,Action callBack=null)
    {
        if (!File.Exists(excelPath))
        {
            MainWindow.LogError($"{excelPath}文件不存在");
            yield break;
        }

        FileInfo fileInfo = new FileInfo(excelPath);
        yield return Excel2Json_File(fileInfo, outPutFloder, callBack);

    }

    public static IEnumerator Excel2Json_File(FileInfo fileInfo, string outPutFloder = null, Action callBack = null)
    {
        MainWindow.Log($"start:{fileInfo.Name}");
        JsonData file_Js = new JsonData();
        try
        {
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheets = package.Workbook.Worksheets;
                bool mutiSheet = false;
                for (int i = sheets.Count - 1; i >= 0; i--)
                {
                    var sheet = package.Workbook.Worksheets[i];
                    if (sheet.Name.StartsWith("_"))
                    {
                        //以'_'开头的名字为注释表格 不转成json
                        continue;
                    }

                    var result = Excel2Json_Sheet(fileInfo.Name, sheet);
                    if (result.Result.GetJsonType() == JsonType.None)
                    {
                        continue;
                    }

                    if (i == 0 && !mutiSheet)
                    {
                        file_Js = result.Result;
                    }
                    else
                    {
                        file_Js[sheet.Name] = result.Result;
                        mutiSheet = true;
                    }
                }
            }
            if (!string.IsNullOrEmpty(outPutFloder))
            {
                if (!Directory.Exists(outPutFloder))
                {
                    Directory.CreateDirectory(outPutFloder);
                }
                string outFilePaht = Path.Combine(outPutFloder, $"{fileInfo.Name.Replace(fileInfo.Extension, "")}.json");
                FileStream fs = new FileStream(outFilePaht, FileMode.Create);
                var data = Encoding.UTF8.GetBytes(Regex.Unescape(file_Js.ToJson()));
                fs.Write(data, 0, data.Length);
                fs.Close();
            }
            MainWindow.Log($"complate:{fileInfo.Name}");
            callBack?.Invoke();
        }
        catch (Exception e)
        {
            MainWindow.LogError($"表格:{fileInfo.Name}读取失败.{e.Message}");
            callBack?.Invoke();
        }

        yield break;
    }

    private static async Task<JsonData> Excel2Json_Sheet(string fileName, ExcelWorksheet sheet)
    {
        JsonData result = new JsonData();
        if (sheet.Dimension.Rows<=2)
        {
            MainWindow.Log($"表格:{fileName},sheet:{sheet.Name},总行数只有两行,跳过该sheet");
            return result;
        }
        int row = 0;
        Dictionary<int, string> colName = new Dictionary<int, string>();
        Dictionary<int, string> colType = new Dictionary<int, string>();
        while (true)
        {
            row++;
            var col_1 = sheet.Cells[row, 1].Value;
            if (col_1 == null)
            {
                if (row<=2)
                {
                    MainWindow.Instance.ShowLogPanel("前两行不能为空,第一行为变量名,第二行为数据类型");
                }
                break;
            }

            switch (row)
            {
                case 1: //变量名称
                    colName = ReadDataHeader(sheet, row);
                    break;
                case 2: //变量类型
                    colType = ReadDataHeader(sheet, row);
                    break;
                default:
                    var row_js = ReadExcelRow(fileName, sheet, row, colName, colType);
                    if (row_js.GetJsonType() != JsonType.None)
                    {
                        if (sheet.Tables.Count > 0)
                        {
                            //数据表
                            //转换成字典样式
                            result[col_1.ToString()] = row_js;
                        }
                        else
                        {
                            //配置信息
                            if (sheet.Dimension.Rows == 3)
                            {
                                result = row_js;
                            }
                            else
                            {
                                result.Add(row_js);
                            }
                        }
                    }

                    break;
            }
        }

        return result;
    }

    private static JsonData ReadExcelRow(string fileName, ExcelWorksheet sheet, int row, Dictionary<int, string> colName, Dictionary<int, string> colType1)
    {
        JsonData result = new JsonData();
        int col = 1;
        var key = sheet.Cells[row, col].Value;
        if (key == null)
        {
            return result;
        }

        while (true)
        {
            if (col > colName.Count)
            {
                break;
            }

            var dataType = colType1[col];
            var value = sheet.Cells[row, col].Value;
            if (value == null)
            {
                // JsonData t = new JsonData();
                // switch (dataType)
                // {
                //     case "Array":
                //         t.SetJsonType(JsonType.Array);
                //         break;
                //     case "string":
                //         t.SetJsonType(JsonType.String);
                //         break;
                //     case "int":
                //         t.SetJsonType(JsonType.Int);
                //         break;
                //     case "double":
                //         t.SetJsonType(JsonType.Double);
                //         break;
                //     case "long":
                //         t.SetJsonType(JsonType.Long);
                //         break;
                //     case "bool":
                //         t.SetJsonType(JsonType.Boolean);
                //         break;
                // }
                // result[colName[col]] = t;
            }
            else
            {
                JsonData temp;
                if (dataType.Contains("Array"))
                {
                    var isSucc = ReadArrayCell(value.ToString(), dataType, out temp);
                    if (!isSucc)
                    {
                        var log = $"表格:{fileName},sheet:{sheet.Name},单元格:{sheet.Cells[row, col]}类型错误";
                        MainWindow.LogError(log);
                        MainWindow.Instance.ShowLogPanel(log);
                    }
                }
                else if(value is DateTime)
                {
                    temp = new JsonData(value.ToString());
                }
                else if(dataType=="int")
                {
                    if (int.TryParse(value.ToString(),out var v))
                    {
                        temp = new JsonData(v);
                    }
                    else
                    {
                        MainWindow.Instance.ShowLogPanel($"表格:{fileName},sheet:{sheet.Name},单元格:{sheet.Cells[row, col]}类型错误");
                        temp = new JsonData(0);
                    }
                }
                else if (dataType == "string")
                {
                    temp = new JsonData(value.ToString());
                }
                else
                {
                    temp = new JsonData(value);
                }

                result[colName[col]] = temp;
            }

            col++;
        }

        return result;
    }

    private static bool ReadArrayCell(string value, string dataType, out JsonData result)
    {
        result = new JsonData();
        result.SetJsonType(JsonType.Array);
        var str_arr = value.Split(',');
        Regex reg = new Regex("<.+?>");
        var math = reg.Match(dataType);
        switch (math.Value)
        {
            case "<string>":
                for (int i = 0; i < str_arr.Length; i++)
                {
                    result.Add(str_arr[i]);
                }

                break;
            case "<int>":
                for (int i = 0; i < str_arr.Length; i++)
                {
                    if (int.TryParse(str_arr[i], out int a))
                    {
                        result.Add(a);
                    }
                    else
                    {
                        return false;
                    }
                }

                break;
            case "<double>":
                for (int i = 0; i < str_arr.Length; i++)
                {
                    if (double.TryParse(str_arr[i], out double a))
                    {
                        result.Add(a);
                    }
                    else
                    {
                        return false;
                    }
                }

                break;
        }

        return true;
    }

    private static Dictionary<int, string> ReadDataHeader(ExcelWorksheet sheet, int row)
    {
        Dictionary<int, string> result = new Dictionary<int, string>();
        int col = 0;
        while (true)
        {
            col++;
            var value = sheet.Cells[row, col].Value;
            if (value == null)
            {
                break;
            }

            result[col] = value.ToString();
        }

        return result;
    }

    public static void Excel2Json_Files(string[] excelPaths, string outPutFloder,Action stepAction=null)
    {
        if (excelPaths == null)
        {
            return;
        }
        for (int i = 0; i < excelPaths.Length; i++)
        {
            string path = excelPaths[i];

            Thread childThread = new Thread( delegate()
            { 
                Excel2Json_File(path, outPutFloder);
               stepAction?.Invoke();
            }){IsBackground = true};
            childThread.Start();

            // yield return Excel2Json_File(path, outPutFloder);
            // stepAction?.Invoke();
        }
    }
}