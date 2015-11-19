using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJsonLib
{
    internal class ColumnInfo
    {
        public string Name;
        public HashSet<JTokenType> TokenTypeSets = new HashSet<JTokenType>();
        public JTokenType Type;
    }

    public static class DataTransform
    {
        public static void Import(Excel.Worksheet activeWorksheet, string path)
        {
            var json = File.ReadAllText(path);
            var datas = JArray.Parse(json);

            // construct column information.
            // if there is an existing column in excel, merge it.

            var colInfos = GetColumeInfosFromExcel(activeWorksheet);

            foreach (JObject data in datas)
            {
                foreach (var prop in data.Properties())
                {
                    var idx = colInfos.FindIndex(c => c.Name == prop.Name);
                    if (idx == -1)
                    {
                        idx = colInfos.Count;
                        colInfos.Add(new ColumnInfo { Name = prop.Name });

                        // when there is new column in json data,
                        // new column will be inserted in excel table.

                        Excel.Range range = activeWorksheet.Columns[idx + 2];
                        range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                                     Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    }
                    colInfos[idx].TokenTypeSets.Add(prop.Value.Type);
                }
            }

            // fill column information at excel table

            for (int i = 0; i < colInfos.Count; i++)
            {
                var tset = colInfos[i].TokenTypeSets;
                if (tset.Count == 1)
                {
                    colInfos[i].Type = tset.First();
                }
                else if (tset.All(t => t == JTokenType.Integer || t == JTokenType.Float))
                {
                    colInfos[i].Type = JTokenType.Float;
                }
                else
                {
                    colInfos[i].Type = JTokenType.None;
                }

                var cell = activeWorksheet.Cells[2, i + 1];
                if (colInfos[i].Type != JTokenType.None)
                {
                    cell.Value = colInfos[i].Name + " : " + colInfos[i].Type;
                }
                else
                {
                    cell.Value = colInfos[i].Name;
                }
            }

            // ensure the column next to last one empty
            // (let this column be a sentinel when export)

            activeWorksheet.Cells[2, colInfos.Count + 1].Clear();

            // make row for each json entity

            var nameToColumnIndexMap = new Dictionary<string, int>();
            for (int ci = 0; ci < colInfos.Count; ci++)
            {
                if (nameToColumnIndexMap.ContainsKey(colInfos[ci].Name) == false)
                    nameToColumnIndexMap[colInfos[ci].Name] = ci;
            }

            var row = 3;
            foreach (JObject data in datas)
            {
                var colValues = new string[colInfos.Count];

                foreach (var prop in data.Properties())
                {
                    var colIdx = nameToColumnIndexMap[prop.Name];
                    var colInfo = colInfos[colIdx];

                    switch (colInfo.Type)
                    {
                        case JTokenType.Integer:
                        case JTokenType.Float:
                        case JTokenType.String:
                            colValues[colIdx] = prop.Value.ToString();
                            break;

                        case JTokenType.Boolean:
                            colValues[colIdx] = (bool)prop.Value ? "true" : "false";
                            break;

                        case JTokenType.Array:
                            colValues[colIdx] = GetStringFromArray((JArray)prop.Value);
                            break;

                        default:
                            colValues[colIdx] = prop.Value.ToString(Formatting.None);
                            break;
                    }
                }

                for (int i = 0; i < colInfos.Count; i++)
                {
                    var cell = activeWorksheet.Cells[row, i + 1];
                    if (string.IsNullOrEmpty(colValues[i]))
                    {
                        if (string.IsNullOrEmpty(cell.Text) == false)
                            cell.Clear();
                    }
                    else if (colInfos[i].Type == JTokenType.Integer || colInfos[i].Type == JTokenType.Float)
                    {
                        if (cell.Value == null || cell.Value.ToString() != colValues[i])
                        {
                            cell.Clear();
                            cell.Value = colValues[i];
                        }
                    }
                    else if (colValues[i] != cell.Text)
                    {
                        cell.Clear();
                        cell.NumberFormat = "@"; // @=Text
                        cell.Value = colValues[i];
                    }
                }

                row += 1;
            }

            // ensure the row next to last one empty
            // (let this row be a sentinel when export)

            for (int i = 0; i < colInfos.Count; i++)
            {
                activeWorksheet.Cells[row, i + 1].Clear();
            }
        }

        private static string GetStringFromArray(JArray array)
        {
            if (array.Count == 0)
                return "";

            bool isOneLine =
                array.Children()
                     .All(t => t.Type == JTokenType.Integer ||
                               t.Type == JTokenType.Float ||
                               t.Type == JTokenType.Boolean);
            if (isOneLine)
            {
                return string.Join(", ",
                                   array.Children().Select(c => c.ToString(Formatting.None)).ToArray());
            }
            else
            {
                return string.Join("\n",
                                   array.Children().Select(c => c.ToString(Formatting.None)).ToArray());
            }
        }

        private static JArray GetArrayFromString(string str)
        {
            return (JArray)JToken.Parse("[" + str.Replace("\n", ",") + "]");
        }

        private static List<ColumnInfo> GetColumeInfosFromExcel(Excel.Worksheet workSheet)
        {
            var colInfos = new List<ColumnInfo>();
            for (int i = 0; i < 1024; i++)
            {
                var cell = workSheet.Cells[2, i + 1];
                var v = cell.Value != null ? (string)cell.Value.ToString() : null;
                if (string.IsNullOrEmpty(v))
                    break;

                var vIdx = v.IndexOf(':');
                if (vIdx != -1)
                {
                    var type = JTokenType.None;
                    Enum.TryParse(v.Substring(vIdx + 1).Trim(), out type);

                    colInfos.Add(new ColumnInfo
                    {
                        Name = v.Substring(0, vIdx).Trim(),
                        Type = type
                    });
                }
                else
                {
                    colInfos.Add(new ColumnInfo
                    {
                        Name = v,
                        Type = JTokenType.None
                    });
                }
            }
            return colInfos;
        }

        public static void Export(Excel.Worksheet activeWorksheet, string path)
        {
            var colInfos = GetColumeInfosFromExcel(activeWorksheet);

            // get data from excel table

            var dataList = new List<Tuple<long, JObject>>();
            for (int i = 0; i < 65536; i++)
            {
                var data = new JObject();
                var dataKey = (long)i;
                var isEmptyRow = true;
                for (int j = 0; j < colInfos.Count; j++)
                {
                    var row = i + 3;
                    var cell = activeWorksheet.Cells[row, j + 1];
                    var v = cell.Value != null ? (string)cell.Value.ToString() : null;
                    if (string.IsNullOrEmpty(v) == false)
                    {
                        isEmptyRow = false;

                        JToken token;
                        switch (colInfos[j].Type)
                        {
                            case JTokenType.Integer:
                                long ival;
                                if (long.TryParse(v, out ival))
                                    token = new JValue(ival);
                                else
                                    throw new Exception(string.Format("[{0} {1}] Integer? {2}", row, j + 1, v));
                                if (j == 0)
                                    dataKey = ival;
                                break;

                            case JTokenType.Float:
                                double fval;
                                if (double.TryParse(v, out fval))
                                    token = new JValue(fval);
                                else
                                    throw new Exception(string.Format("[{0} {1}] Float? {2}", row, j + 1, v));
                                break;

                            case JTokenType.String:
                                token = new JValue(v);
                                break;

                            case JTokenType.Boolean:
                                if (string.Compare(v, "true", StringComparison.InvariantCultureIgnoreCase) == 0)
                                    token = new JValue(true);
                                else if (string.Compare(v, "false", StringComparison.InvariantCultureIgnoreCase) == 0)
                                    token = new JValue(false);
                                else
                                    throw new Exception(string.Format("[{0} {1}] Boolean? {2}", row, j + 1, v));
                                break;

                            case JTokenType.Array:
                                try
                                {
                                    token = GetArrayFromString(v);
                                }
                                catch (Exception e)
                                {
                                    throw new Exception(string.Format("[{0} {1}] Json? {2}\n{3}", row, j + 1, v, e));
                                }
                                break;

                            default:
                                try
                                {
                                    token = JToken.Parse(v);
                                }
                                catch (Exception e)
                                {
                                    throw new Exception(string.Format("[{0} {1}] Json? {2}\n{3}", row, j + 1, v, e));
                                }
                                break;
                        }
                        data[colInfos[j].Name] = token;
                    }
                }
                if (isEmptyRow)
                    break;

                dataList.Add(Tuple.Create(dataKey, data));
            }

            // sort data by key

            var datas = new JArray();
            foreach (var entity in dataList.OrderBy(row => row.Item1))
                datas.Add(entity.Item2);

            // save to json

            var json = JsonConvert.SerializeObject(
                datas,
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    DefaultValueHandling = DefaultValueHandling.Ignore
                });
            var prettyJson = JsonUtility.PrettifyJson(json);
            File.WriteAllText(path, prettyJson, new UTF8Encoding(true));
        }
    }
}
