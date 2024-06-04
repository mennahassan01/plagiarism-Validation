using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using ExcelDataReader;
using System.Text;

using System.IO;
using System.Diagnostics;

using OfficeOpenXml;
using System.ComponentModel;

public class ProblemClass
{
    public static void RequiredFunction(List<MatchingPair> edges, List<KeyValuePair<float, KeyValuePair<List<string>, int>>> stat)
    {
        Dictionary<string, List<KeyValuePair<string, float>>> graph = new Dictionary<string, List<KeyValuePair<string, float>>>();

        foreach (var edge in edges)
        {
            string id1 = edge.ID1;
            string id2 = edge.ID2;
            float sim1 = float.Parse(edge.Sim1);
            float sim2 = float.Parse(edge.Sim2);
            float weight = (float)Math.Round((sim1 + sim2) / 2, 1);


            if (!graph.ContainsKey(id1))
            {
                graph[id1] = new List<KeyValuePair<string, float>>();
            }
            graph[id1].Add(new KeyValuePair<string, float>(id2, weight));

            if (!graph.ContainsKey(id2))
            {
                graph[id2] = new List<KeyValuePair<string, float>>();
            }
            graph[id2].Add(new KeyValuePair<string, float>(id1, weight));
        }

        HashSet<string> visited = new HashSet<string>();

        foreach (var node in graph.Keys)
        {
            if (!visited.Contains(node))
            {
                List<string> set = new List<string>();
                float sum = 0f;
                int count = 0;
                DFS(node, graph, visited, set, ref sum, ref count);

                float avg = (float)Math.Round((sum / count), 1);
                int size = set.Count;
                stat.Add(new KeyValuePair<float, KeyValuePair<List<string>, int>>(avg, new KeyValuePair<List<string>, int>(set, size)));

            }
        }
    }

    private static void DFS(string node, Dictionary<string, List<KeyValuePair<string, float>>> graph, HashSet<string> visited, List<string> set, ref float sum, ref int count)
    {
        visited.Add(node);
        set.Add(node);

        if (graph.ContainsKey(node))
        {
            foreach (var neighbor in graph[node])
            {
                float weight = neighbor.Value;
                sum += weight;
                count++;

                string neighborID = neighbor.Key;
                if (!visited.Contains(neighborID))
                {
                    DFS(neighborID, graph, visited, set, ref sum, ref count);
                }
            }
        }
    }

    static void SaveComponentStatsToExcel(List<KeyValuePair<float, KeyValuePair<List<string>, int>>> stats)
    {
        var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Stat_File.xlsx");


        using (ExcelPackage excel = new ExcelPackage())
        {
            var ws = excel.Workbook.Worksheets.Add("Statistics");
            ws.Cells["A1"].Value = "Component Index";
            ws.Cells["B1"].Value = "Vertices";
            ws.Cells["C1"].Value = "Average Similarity";
            ws.Cells["D1"].Value = "Component Count";

            int index = 1;
            int row = 2;
            foreach (var stat in stats)
            {
                ws.Cells[row, 1].Value = index;
                ws.Cells[row, 2].Value = string.Join(", ", stat.Value.Key); // Convert List<string> to string
                ws.Cells[row, 3].Value = Math.Round(stat.Key, 1); ;
                ws.Cells[row, 4].Value = stat.Value.Value;
                row++;
                index++;
            }

            FileInfo excelFile = new FileInfo(filePath);
            excel.SaveAs(excelFile);
        }

        Console.WriteLine($"Component statistics saved to {filePath}");


    }
    static void Save_MST_IN_Excel(Dictionary<HashSet<string>, List<Tuple<string, string, Double, string, string, object>>> MST)
    {
        var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), " MST_File.xlsx");


        using (ExcelPackage excel = new ExcelPackage())
        {
            var ws = excel.Workbook.Worksheets.Add("MST");
            ws.Cells["A1"].Value = "File 1";
            ws.Cells["B1"].Value = "File 2";
            ws.Cells["C1"].Value = "Lines Matches";





            int row = 2;
            foreach (var kvp in MST)
            {
                foreach (var tuple in kvp.Value)
                {
                    if (tuple.Item6 != null)
                    {
                        ws.Cells[row, 1].Hyperlink = new ExcelHyperLink(tuple.Item6.ToString());
                        ws.Cells[row, 2].Hyperlink = new ExcelHyperLink(tuple.Item6.ToString());
                    }
                    ws.Cells[row, 1].Value = tuple.Item1.ToString();
                    ws.Cells[row, 2].Value = tuple.Item2;
                    ws.Cells[row, 3].Value = tuple.Item3;

                    ws.Cells[row, 1].Style.Font.UnderLine = true;
                    ws.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                    ws.Cells[row, 2].Style.Font.UnderLine = true;
                    ws.Cells[row, 2].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    row++;
                }


            }
            ws.Cells.AutoFitColumns();
            FileInfo excelFile = new FileInfo(filePath);
            excel.SaveAs(excelFile);
        }

        Console.WriteLine($"MST File saved to {filePath}");


    }
    public static void Main(string[] args)
    {
        var stat_time = new Stopwatch();
        var MST_time = new Stopwatch();
        var All_time = new Stopwatch();
        var read_time = new Stopwatch();
        var write_time_in_stat = new Stopwatch();
        All_time.Start();
        read_time.Start();
        String filePath = "Complete/Easy/1-Input.xlsx";

        List<MatchingPair> input = ReadMatchingPairs(filePath);
        read_time.Stop();

        List<KeyValuePair<float, KeyValuePair<List<string>, int>>> stat = new List<KeyValuePair<float, KeyValuePair<List<string>, int>>>();
        stat_time.Start();
        RequiredFunction(input, stat);

        stat.Sort((x, y) => y.Key.CompareTo(x.Key));
        foreach (var s in stat)
        {
            float key = s.Key;
            s.Value.Key.Sort((x, y) => Convert.ToInt32(x).CompareTo(Convert.ToInt32(y)));

        }
        stat_time.Stop();
        //foreach (var kvp in stat)
        //{
        //    float key = kvp.Key;
        //    int size = kvp.Value.Value;
        //    Console.Write($"{key} -> ");
        //    Console.WriteLine();
        //    foreach (var value in kvp.Value.Key)
        //    {
        //        Console.Write(value + ", ");

        //    }

        //    Console.WriteLine(" -> " + size);
        //    Console.WriteLine();
        //}




        //  Console.ReadLine();
        Dictionary<HashSet<string>, List<Tuple<string, string, Double, string, string, object>>> MST = new Dictionary<HashSet<string>, List<Tuple<string, string, double, string, string, object>>>();


        MST_time.Start();
        List<MatchingPair> sortedList = input.OrderByDescending(pair => pair.Similarity).ThenByDescending(pair => pair.LinesMatched).ToList();
        foreach (var kvp in stat)
        {
            HashSet<string> hashvalues = new HashSet<string>();
            foreach (var value in kvp.Value.Key)
            {
                hashvalues.Add(value);
            }
            MST[hashvalues] = new List<Tuple<string, string, double, string, string, object>>();
        }
        Dictionary<string, string> parentsss = new Dictionary<string, string>();
        foreach (var s in stat)
        {
            foreach (var i in s.Value.Key)
            {
                parentsss[i] = i;
            }
        }

        foreach (var row in sortedList)
        {

            var edge = Tuple.Create(row.File1, row.File2, row.LinesMatched, row.ID1, row.ID2, row.lnk);
            merge(edge, parentsss, MST);
        }

        foreach (var kvp in MST)
        {
            kvp.Value.Sort((x, y) => y.Item3.CompareTo(x.Item3));
        }
        //  int count = 0;
        //foreach (var kvp in MST)
        //{
        //    foreach (var tuple in kvp.Value)
        //    {
        //        Console.WriteLine($"{tuple.Item1}, {tuple.Item2}, {tuple.Item3}");
        //        count++;
        //    }
        //}
        //Console.WriteLine(count);


        Save_MST_IN_Excel(MST);
        MST_time.Stop();
        write_time_in_stat.Start();
        SaveComponentStatsToExcel(stat);
        write_time_in_stat.Stop();
        All_time.Stop();
        var tim = stat_time.ElapsedMilliseconds + write_time_in_stat.ElapsedMilliseconds;
        Console.WriteLine("time of read only :" + read_time.ElapsedMilliseconds + "ms");
        // Console.WriteLine("time of stat only :" + stat_time.ElapsedMilliseconds + "ms");
        Console.WriteLine("time of stat + save :" + tim + "ms");
        Console.WriteLine("time of MST + saving in excel :" + MST_time.ElapsedMilliseconds + "ms");
        Console.WriteLine("time of all : " + All_time.ElapsedMilliseconds + "ms");
        Console.ReadLine();
    }


    static string findParent(string i, Dictionary<string, string> parentsss)
    {
        if (parentsss[i] == i)
            return i;
        return parentsss[i] = findParent(parentsss[i], parentsss);
    }
    static void merge(Tuple<string, string, Double, string, string, object> edge, Dictionary<string, string> parentsss, Dictionary<HashSet<string>, List<Tuple<string, string, Double, string, string, object>>> MST)
    {
        string id1 = edge.Item4;
        string id2 = edge.Item5;

        string parentOne = findParent(id1, parentsss);
        string parentTwo = findParent(id2, parentsss);
        if (parentOne != parentTwo)
        {
            parentsss[parentTwo] = parentOne;
            foreach (var key in MST.Keys)
            {
                if (key.Contains(id1) || key.Contains(id2))
                {
                    MST[key].Add(edge);
                    break;
                }
            }
        }
    }
    static string GetLastDigits(string file)
    {
        string similarity = "";
        for (int i = file.Length - 1; i >= 0 && similarity.Length < 3; i--)
        {
            if (file[i] == '(') break;
            if (char.IsDigit(file[i]))
            {
                similarity = file[i] + similarity;
            }
        }
        return similarity;
    }
    public class MatchingPair
    {
        public string File1 { get; }
        public string File2 { get; }
        public double Similarity { get; }
        public double LinesMatched { get; }

        public string Sim1 = "";
        public string Sim2 = "";

        public string ID1 = "";
        public string ID2 = "";
        public object lnk { get; }
        public MatchingPair(string file1, string file2, double similarity, double linesMatched, string sim1, string sim2, string id1, string id2, object hyplink)
        {
            File1 = file1;
            File2 = file2;
            Similarity = similarity;
            LinesMatched = linesMatched;
            Sim1 = sim1;
            Sim2 = sim2;
            ID1 = id1;
            ID2 = id2;
            lnk = hyplink;
        }
    }
    public static List<MatchingPair> ReadMatchingPairs(string filePath)
    {

        List<MatchingPair> matchingPairs = new List<MatchingPair>();

 
        using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
        {

            if (package.Workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("No worksheets found in the Excel file.");
                return matchingPairs;
            }


            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

            if (worksheet == null)
            {
                Console.WriteLine("No worksheet found in the Excel file.");
                return matchingPairs;
            }
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                if (worksheet.Cells[row, 1].Value == null || worksheet.Cells[row, 2].Value == null)
                    break;
                string file1 = worksheet.Cells[row, 1].Value.ToString(); // First column
                string file2 = worksheet.Cells[row, 2].Value.ToString(); // Second column
                double linesMatched = (double)worksheet.Cells[row, 3].Value;

                string id1 = "";
                string id2 = "";
                int count = 0;
                foreach (char i in file1)
                {
                    if (Char.IsDigit(i))
                    {
                        count++;
                        id1 += i;
                    }
                    if (i == '/' && count > 0)
                    {
                        break;
                    }
                }
                count = 0;
                foreach (char i in file2)
                {
                    if (Char.IsDigit(i))
                    {
                        count++;
                        id2 += i;
                    }
                    if (i == '/' && count > 0)
                    {
                        break;
                    }
                }
                string sim1 = GetLastDigits(file1);
                string sim2 = GetLastDigits(file2);

                double similarity = Math.Max(int.Parse(sim1), int.Parse(sim2));

                object hyper = worksheet.Cells[row, 1].Hyperlink;

                matchingPairs.Add(new MatchingPair(file1, file2, similarity, linesMatched, sim1, sim2, id1, id2, hyper));

            }
        }
        return matchingPairs;
    }


}