using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileParsing
{
    class EOT83
    {
        public static void Parse()
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
                try
                {
                    Stopwatch timer = new Stopwatch();
                    timer.Reset();
                    timer.Start();
                    var a = File.ReadAllLines(file).ToList();
                    List<string[]> e = new List<string[]>();
                    e.Add(a[8].Replace("\t", "").Split('='));
                    e.Add(a[9].Replace("\t", "").Split('='));
                    e.Add(a[10].Replace("\t", "").Split('='));
                    a.RemoveRange(0, 15);

                    foreach (string[] data in e)
                    {
                        a[0] += "\t" + data[0].Trim();
                        for (int i = 1; i < a.Count; i++)
                            a[i] += "\t" + data[1].Trim();
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
                        data1.Add(string.Format("\"{0}\",", line.Replace("\t", "<<DIV>>").Replace("\"", "") + "<<EOL>>"));
                    }
                    File.AppendAllLines(filePath, data1);
                    Console.WriteLine("Parsed : " + Path.GetFileName(file) + " Time taken: " + timer.Elapsed.ToString(@"m\:ss\.ff"));
                    GC.Collect();
                }
                catch (Exception _)
                {
                    GC.Collect();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error in parsing : " + Path.GetFileName(file));
                    Console.WriteLine("Error Message : " + _.Message ?? _.InnerException.ToString());
                    Console.ResetColor();
                    continue;
                }

            }
            Console.ReadLine();
        }
    }
}
