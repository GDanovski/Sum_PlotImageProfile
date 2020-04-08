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
            int RefIndex = 0;
            Dictionary<string, Data> inputFiles = new Dictionary<string, Data>();
            do
            {
                Console.WriteLine("input dir:");
                dir = Console.ReadLine();
                if (!dir.EndsWith(@"\")) dir += @"\";
            }
            while (!Directory.Exists(dir));

            do
            {
                Console.WriteLine("Ref index:");
            }
            while (!int.TryParse(Console.ReadLine(), out RefIndex));

            GetAllFiles(".txt",dir,inputFiles);
            ProcessFiles(inputFiles,RefIndex);
            string[] results = CalculateResults(inputFiles);

            File.WriteAllLines(dir + "results_FirsFocus.txt", results);

            results = CalculateResultsSecondFocus(inputFiles);
            File.WriteAllLines(dir + "results_SecondFocus.txt", results);

            Console.WriteLine("Done!");

            Console.ReadLine();
        }
        private static void GetAllFiles(string ext, string dir, Dictionary<string, Data> inputFiles)
        {
            foreach (string file in Directory.GetFiles(dir))
                if (file.EndsWith(ext) && !file.EndsWith("Focus.txt"))
                    inputFiles.Add(file, new Data());

            foreach (string folder in Directory.GetDirectories(dir))
                GetAllFiles(ext, folder, inputFiles);
        }
        private static void ProcessFiles(Dictionary<string, Data> inputFiles, int Refindex)
        {
            Parallel.ForEach(inputFiles, (file) => {
                file.Value.FileName = Path.GetFileNameWithoutExtension(file.Key);                
                file.Value.SetInputValues(File.ReadAllLines(file.Key),Refindex);
            });
        }
        private static string[] CalculateResults(Dictionary<string, Data> inputFiles)
        {
            int MaxRows = FindMaxRows(inputFiles);
            string[] output = new string[MaxRows + 1];

            int column, row;
            List<string> cur = new List<string>();
            cur.Add("Index");

            for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
            {
                foreach (var kvp in inputFiles)
                    cur.Add(kvp.Value.GetTitle(column) + "_" + kvp.Value.FileName);

                cur.Add("Avg_" + inputFiles.ElementAt(0).Value.GetTitle(column));
                cur.Add("StDev_" + inputFiles.ElementAt(0).Value.GetTitle(column));
            }

            output[0] = string.Join("\t", cur);
            double val;
            string ChartRange = "$A:$A";
            bool addToExcell = true;
            for (row = 0; row < MaxRows; row++)
            {
                int first = 1;
                int last = inputFiles.Count;
                int totalLength = last + 2;

                cur.Clear();
                cur.Add(row.ToString());

                for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
                {
                    foreach (var kvp in inputFiles)
                    {
                        val = kvp.Value.GetValue(column, row);
                        cur.Add(val != 0 ? val.ToString() : "");
                    }

                    string Avg = GetExcelCommand("AVERAGE", first, last, row);
                    if(addToExcell)
                        ChartRange += ",$" + GetColumnName(last + 1) + ":$" + GetColumnName(last + 1);
                    string StDev = GetExcelCommand("STDEV.S", first, last, row);

                    first += totalLength;
                    last += totalLength;

                    cur.Add(Avg);
                    cur.Add(StDev);

                }
                addToExcell = false;
                output[row + 1] = string.Join("\t", cur);
            }
            ChartRange = ChartRange.Substring(1);

           // Console.WriteLine(GetMacro(ChartRange));
            return output;
        }
        private static string[] CalculateResultsSecondFocus(Dictionary<string, Data> inputFiles)
        {
            int MaxRows = FindMaxRows(inputFiles);
            string[] output = new string[MaxRows + 1];

            int column, row;
            List<string> cur = new List<string>();
            cur.Add("Index");
            for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
            {
                foreach (var kvp in inputFiles)
                    cur.Add(kvp.Value.GetTitle(column) + "_" + kvp.Value.FileName);

                cur.Add("Avg_" + inputFiles.ElementAt(0).Value.GetTitle(column));
                cur.Add("StDev_" + inputFiles.ElementAt(0).Value.GetTitle(column));
            }

            output[0] = string.Join("\t", cur);
            double val;
            string ChartRange = "$A:$A";
            bool addToExcell = true;
            for (row = 0; row < MaxRows; row++)
            {
                int first = 1;
                int last = inputFiles.Count;
                int totalLength = last + 2;

                cur.Clear();
                cur.Add(row.ToString());

                for (column = 0; column < inputFiles.ElementAt(0).Value.GetColumns; column++)
                {
                    foreach (var kvp in inputFiles)
                    {
                        val = kvp.Value.GetValueSecondFocus(column, row);
                        cur.Add(val != 0 ? val.ToString() : "");
                    }

                    string Avg = GetExcelCommand("AVERAGE", first, last, row);
                    string StDev = GetExcelCommand("STDEV.S", first, last, row);
                    if(addToExcell)
                        ChartRange += ",$" + GetColumnName(last + 1) + ":$" + GetColumnName(last + 1);
                    first += totalLength;
                    last += totalLength;

                    cur.Add(Avg);
                    cur.Add(StDev);

                }
                addToExcell = false;
                output[row + 1] = string.Join("\t", cur);
            }
            ChartRange = ChartRange.Substring(1);
            Console.WriteLine(GetMacro(ChartRange));
            return output;
        }
        private static string GetMacro(string range)
        {
            return (@"
-----------------------------------------
Excel macro
-----------------------------------------

    Range('" + range.Replace("$","") + @"').Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range('" + range + @"')

-----------------------------------------

").Replace("'","\"");
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
        //private int maxIndex = 0;
        private string[] titles;
        private double[,] values;//[column,row]
        private double[,] valuesSecFocus;//[column,row]

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
        public double GetValueSecondFocus(int column, int row)
        {
            if (column < GetColumns && row < GetRows)
                return valuesSecFocus[column, row];
            else
                return 0d;
        }
        public string GetTitle(int column)
        {
            return titles[column];
        }
        public void SetInputValues(string[] input, int RefIndex)
        {
            if (input == null || input.Length < 2) return;

            double MaxValue = double.MinValue;
           int maxIndex = 0;

            titles = input[0].Split(new string[] { "\t" }, StringSplitOptions.None);

           var newValues = new double[GetColumns, input.Length - 1];
            string[] rowString;
            //extract double values and find max value index
            for (int row = 0; row < input.Length - 1; row++)
            {
                rowString = input[row+1].Split(new string[] { "\t" }, StringSplitOptions.None);

                for (int column = 0; column < GetColumns; column++)
                    double.TryParse(rowString[column], out newValues[column, row]);

                if(newValues[RefIndex, row] > MaxValue)
                {
                    MaxValue = newValues[RefIndex, row];
                    maxIndex = row;
                }
            }
            //recalculate matrix
            if (maxIndex < (double)newValues.GetLength(1) / 2d)//copy with new index
            {
                this.values = new double[GetColumns, newValues.GetLength(1) - maxIndex];

                for (int row = maxIndex, newRow = 0; row < newValues.GetLength(1); row++,newRow++)
                    for (int column = 0; column < GetColumns; column++)
                        this.values[column, newRow] = newValues[column, row];
            }
            else//revert and copy
            {
                this.values = new double[GetColumns, maxIndex + 1];

                for (int row = maxIndex, newRow = 0; row >= 0; row--, newRow++)
                    for (int column = 0; column < GetColumns; column++)
                        this.values[column, newRow] = newValues[column, row];
            }

            CalculateSecFoci(RefIndex);
        }
        public void CalculateSecFoci(int RefIndex)
        {
            int col, row, newRow, Half,maxIndex;
            double Max;
            this.valuesSecFocus = new double[GetColumns, GetRows];

            //find half
            col = GetColumns / 2 + RefIndex;
            if (col >= GetColumns) col = GetColumns - 1;
            Half = 0;
            for (row = GetRows - 1; row > 0; row--)
                if (GetValue(col, row) != 0d)
                {
                    Half = (int)(row / 2);
                    break;
                }

            for (col = 0; col < GetColumns; col++)
            {                
                //find max index
                Max = double.MinValue;
                maxIndex = 0;
                for (row = Half;row<GetRows;row++)
                    if(GetValue(col, row) > Max)
                    {
                        Max = GetValue(col, row);
                        maxIndex = row;
                    }
                //copy reversed
                
                for (row = maxIndex, newRow = 0; row >= 0; row--, newRow++)
                        this.valuesSecFocus[col, newRow] = this.values[col, row];
            }
        }
    }
}
