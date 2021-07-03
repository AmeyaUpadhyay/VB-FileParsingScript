using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileParsing
{
    class EOT77
    {
        static void Parse()
        {
            int count = 0;
            Console.WriteLine("Please enter the Path to read JavaScript Files.!");

            readPath: var pathToRead = Console.ReadLine();

            if (string.IsNullOrEmpty(pathToRead) || !Directory.Exists(pathToRead))
            {
                count++;
                if (count > 5)
                {
                    Console.WriteLine("TOO MANY INVALID PATHS\r\n");
                    Environment.Exit(0);
                }
                Console.WriteLine("The path is not valid, Please re-enter the path\r\n");
                goto readPath;
            }

            var allFiles = Directory.GetFiles(pathToRead, @"*.xls").ToList();

            foreach (var file in allFiles)
            {
                Stopwatch timer = new Stopwatch();
                timer.Reset();
                timer.Start();
                try
                {
                    var a = File.ReadAllLines(file).ToList();
                    List<string[]> e = new List<string[]>();
                    e.Add(a[8].Replace("\t", "").Split('='));
                    e.Add(a[9].Replace("\t", "").Split('='));
                    e.Add(a[10].Replace("\t", "").Split('='));
                    e.Add(a[11].Replace("\t", "").Split('='));
                    e.Add(a[12].Replace("\t", "").Split('='));
                    e.Add(a[13].Replace("\t", "").Split('='));
                    e.Add(a[14].Replace("\t", "").Split('='));
                    a.RemoveRange(0, 19);
                    List<string> columnsToBeAdded = new List<string>()
                    {
                        "Ledger",
                        "Booking Currency",
                        "Report Account",
                        "Report Period",
                        "Period End Date",
                        "Run At Date"
                    };
                    Dictionary<string, List<string>> unsortedColumns = new Dictionary<string, List<string>>();
                    foreach (string[] data in e)
                    {
                        a[0] += "\t" + data[0].Trim();
                        for (int i = 0; i < a.Count; i++)
                        {
                            if (a[i].StartsWith(columnsToBeAdded[0]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[0]))
                                    unsortedColumns[columnsToBeAdded[0]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[0], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (a[i].StartsWith(columnsToBeAdded[1]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[1]))
                                    unsortedColumns[columnsToBeAdded[1]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[1], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (a[i].StartsWith(columnsToBeAdded[2]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[2]))
                                    unsortedColumns[columnsToBeAdded[2]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[2], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (a[i].StartsWith(columnsToBeAdded[3]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[3]))
                                    unsortedColumns[columnsToBeAdded[3]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[3], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (a[i].StartsWith(columnsToBeAdded[4]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[4]))
                                    unsortedColumns[columnsToBeAdded[4]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[4], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (a[i].StartsWith(columnsToBeAdded[5]))
                            {
                                if (unsortedColumns.ContainsKey(columnsToBeAdded[5]))
                                    unsortedColumns[columnsToBeAdded[5]].Add(a[i].Replace("\t", "").Split(":")[1]);
                                else
                                    unsortedColumns.Add(columnsToBeAdded[5], new List<string>() { a[i].Replace("\t", "").Split(":")[1] });
                                a.RemoveAt(i);
                            }
                            if (string.IsNullOrEmpty(a[i].Replace("\t", "")))
                                a.RemoveAt(i);
                            else
                                a[i] += "\t" + data[1].Trim();
                        }
                    }

                    foreach (KeyValuePair<string, List<string>> keyValuePair in unsortedColumns)
                        a[0] += "\t" + keyValuePair.Key.Trim();

                    for (int i = 1; i < unsortedColumns.Keys.Count; i++)
                    {
                        List<string> currentDataToAdd = unsortedColumns.ElementAt(i).Value;
                        for (int j = 0; j < currentDataToAdd.Count; j++)
                            a[j] += "\t" + currentDataToAdd[j];
                    }
                    var parsedCSVPath = file.Split(Path.GetFileName(file))[0] + "ParsedCSV\\";
                    if (!Directory.Exists(parsedCSVPath)) Directory.CreateDirectory(parsedCSVPath);
                    string filePath = parsedCSVPath + Path.GetFileName(file).Replace(".xls", ".csv");
                    if (!File.Exists(filePath))
                    { var stream = File.Create(filePath); stream.Close(); }
                    else
                    {
                        File.Delete(filePath);
                        var stream = File.Create(filePath);
                        stream.Close();
                    }
                    var data1 = new List<string>();
                    foreach (string line in a)
                    {
                        data1.Add(String.Format("\"{0}\",", line.Replace("\t", "<<DIV>>").Replace("\"", "") + "<<EOL>>"));
                    }
                    File.AppendAllLines(filePath, data1);
                    Console.WriteLine("Parsed : " + Path.GetFileName(file) + " Time taken: " + timer.Elapsed.ToString(@"m\:ss\.ff"));
                    GC.Collect();
                }
                catch (Exception _)
                {
                    GC.Collect();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error in parsing : " + file);
                    Console.WriteLine("Error Message : " + _.Message ?? _.InnerException.ToString());
                    Console.ResetColor();
                    continue;
                }
            }

            Console.ReadLine();

        }
    }
}
