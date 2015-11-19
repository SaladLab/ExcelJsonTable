using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelJsonLib
{
    // This module translates json
    // from:
    //   "key": 
    //   [
    //      1,
    //      2,
    //      3
    //   ]
    // to:
    //   "key": [ 1, 2, 3 ]
    public static class JsonUtility
    {
        public const int OneLineThreshold = 40;

        public static string PrettifyJson(string json)
        {
            var reIndent = new Regex(@"\r*\n\s+", RegexOptions.Multiline);

            var sb = new StringBuilder();
            var nextCopyIndex = 0;
            foreach (var i in JsonNonStringIndexes(json, 0, json.Length))
            {
                if (i < nextCopyIndex)
                    continue;

                var ic = json[i];
                if (ic == '[')
                {
                    // get closing ] position (endIndex)
                    var endIndex = 0;
                    var depth = 1;
                    foreach (var j in JsonNonStringIndexes(json, i + 1, json.Length - i - 1))
                    {
                        var jc = json[j];
                        if (jc == '[')
                        {
                            depth += 1;
                        }
                        else if (jc == ']')
                        {
                            depth -= 1;
                            if (depth == 0)
                            {
                                endIndex = j;
                                break;
                            }
                        }
                        else if (jc == '{')
                        {
                            endIndex = 0;
                            break;
                        }
                    }

                    // measure the length of between [ ]
                    if (endIndex > 0)
                    {
                        var str = json.Substring(i, endIndex - i + 1);
                        var target = reIndent.Replace(str, " ");
                        if (target.Length < str.Length && target.Length < OneLineThreshold)
                        {
                            sb.Append(json.Substring(nextCopyIndex, i - nextCopyIndex));
                            sb.Append(target);
                            nextCopyIndex = endIndex + 1;
                        }
                    }
                }
            }
            if (nextCopyIndex < json.Length)
                sb.Append(json.Substring(nextCopyIndex, json.Length - nextCopyIndex));

            return sb.ToString();
        }

        private static IEnumerable<int> JsonNonStringIndexes(string json, int startIndex, int count)
        {
            var end = startIndex + count;
            var inString = false;
            for (int i = startIndex; i < end; i++)
            {
                var c = json[i];

                if (c == '"')
                {
                    if (inString)
                    {
                        var escaped = false;
                        for (int j = i - 1; j >= 0; j--)
                        {
                            if (json[j] == '\\')
                                escaped = !escaped;
                            else
                                break;
                        }
                        if (escaped == false)
                        {
                            inString = false;
                            yield return i;
                        }
                    }
                    else
                    {
                        inString = true;
                        yield return i;
                    }
                }
                else
                {
                    if (inString == false)
                        yield return i;
                }
            }
        }
    }
}
