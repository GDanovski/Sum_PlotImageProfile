using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Sum_PlotimageProfile
{
    class Program
    {
        static void Main(string[] args)
        {
            string dir;
            Dictionary<string, Data> inputFiles = new Dictionary<string, Data>();
            do
            {
                Console.WriteLine("input dir:");
                dir = Console.ReadLine();
                if (!dir.EndsWith(@"\")) dir += @"\";
            }
            while (!Directory.Exists(dir));

            GetAllFiles(".txt",dir,inputFiles);
            ProcessFiles(inputFiles);
            string[] results = CalculateResults(inputFiles);

            File.WriteAllLines(dir + "results.txt", results);
            Console.WriteLine("Done!");

            Console.ReadLine();
        }
        private static void GetAllFiles(string ext, string dir, Dictionary<string, Data> inputFiles)
        {
            foreach (string file in Directory.GetFiles(dir))
                if (file.EndsWith(ext))
                    inputFiles.Add(file, new Data());

            foreach (string folder in Directory.GetDirectories(dir))
                GetAllFiles(ext, folder, inputFiles);
        }
        private static void ProcessFiles(Dictionary<string, Data> inputFiles)
        {
            Parallel.ForEach(inputFiles, (file) => {
                file.Value.FileName = Path.GetFileNameWithoutExtension(file.Key);                
                file.Value.SetInputValues(File.ReadAllLines(file.Key));
            });
        }
        private static string[] CalculateResults(Dictionary<string, Data> inputFiles)
        {
            int MaxRows = FindMaxRows(inputFiles);
            string[] output = new string[MaxRows + 1];

            int column, row;
            List<string> cur = new List<string>();


            for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
            {
                foreach (var kvp in inputFiles)
                    cur.Add(kvp.Value.GetTitle(column) + "_" + kvp.Value.FileName);

                cur.Add("Avg_" + inputFiles.ElementAt(0).Value.GetTitle(column));
                cur.Add("StDev_" + inputFiles.ElementAt(0).Value.GetTitle(column));
            }

            output[0] = string.Join("\t", cur);
            
            

            for (row = 0; row < MaxRows; row++)
            {
                int first = 0;
                int last = inputFiles.Count - 1;
                int totalLength = last + 2;

                cur.Clear();

                for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
                {  
                    foreach (var kvp in inputFiles)
                        cur.Add(kvp.Value.GetValue(column, row).ToString());

                    string Avg = GetExcelCommand("AVERAGE", first, last, row);
                    string StDev = GetExcelCommand("STDEV.S", first, last, row);

                    first += totalLength + 1;
                    last += totalLength + 1;

                    cur.Add(Avg);
                    cur.Add(StDev);

                }

                output[row+1] = string.Join("\t", cur);
            }

            return output;
        }
        private static int FindMaxRows(Dictionary<string, Data> inputFiles)
        {
            int MaxRows = int.MinValue;

            foreach (var kvp in inputFiles)
                if (kvp.Value.GetRows > MaxRows)
                    MaxRows = kvp.Value.GetRows;

            return MaxRows;
        }
        private static string GetExcelCommand( string command, int first, int last, int row)
        {
            string prefix = "=" + command + "(" + GetColumnName(first);
            string suffix = ":" + GetColumnName(last);
            
                return prefix + (row + 2).ToString() + suffix + (row + 2).ToString() + ")";            
        }
        private static string GetColumnName(int index) // zero-based
        {
            const byte BASE = 'Z' - 'A' + 1;
            string name = String.Empty;

            do
            {
                name = Convert.ToChar('A' + index % BASE) + name;
                index = index / BASE - 1;
            }
            while (index >= 0);

            return name;
        }
    }
    class Data
    {
        private string fileName;
        private int maxIndex = 0;
        private string[] titles;
        private double[,] values;//[column,row]
        public string FileName
        {
            set
            {
                this.fileName = value;
            }
            get
            {
                return this.fileName;
            }
        }
        public int GetRows
        {
            get
            {
                return values.GetLength(1);
            }
        }
        public int GetColumns
        {
            get
            {
                return titles.Length-1;
            }
        }
        public double GetValue (int column, int row)
        {
            if (column < GetColumns && row < GetRows)
                return values[column, row];
            else
                return 0d;
        }
        public string GetTitle(int column)
        {
            return titles[column];
        }
        public void SetInputValues(string[] input)
        {
            if (input == null || input.Length < 2) return;

            double MaxValue = double.MinValue;
            

            titles = input[0].Split(new string[] { "\t" }, StringSplitOptions.None);

           var newValues = new double[GetColumns, input.Length - 1];
            string[] rowString;
            //extract double values and find max value index
            for (int row = 0; row < input.Length - 1; row++)
            {
                rowString = input[row+1].Split(new string[] { "\t" }, StringSplitOptions.None);

                for (int column = 0; column < GetColumns; column++)
                    double.TryParse(rowString[column], out newValues[column, row]);

                if(newValues[5, row] > MaxValue)
                {
                    MaxValue = newValues[5, row];
                    this.maxIndex = row;
                }
            }
            //recalculate matrix
            if (maxIndex < (double)newValues.GetLength(1) / 2d)//copy with new index
            {
                this.values = new double[GetColumns, newValues.GetLength(1) - this.maxIndex];

                for (int row = this.maxIndex, newRow = 0; row < newValues.GetLength(1); row++,newRow++)
                    for (int column = 0; column < GetColumns; column++)
                        this.values[column, newRow] = newValues[column, row];
            }
            else//revert and copy
            {
                this.values = new double[GetColumns, this.maxIndex + 1];

                for (int row = this.maxIndex, newRow = 0; row >= 0; row--, newRow++)
                    for (int column = 0; column < GetColumns; column++)
                        this.values[column, newRow] = newValues[column, row];
            }

        }
    }
}
