using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\***\exhibitA-input.csv");
                Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range xlRange = excelWorksheet.UsedRange;
                HashSet<MusicStreaming> streamList = new HashSet<MusicStreaming>();

                DateTime end;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                MusicStreaming ms = null;
                DateTime start = DateTime.Now;

                for (int i = 2; i <= rowCount; i++)                                           //starting from second line, not to take header line
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            ms = new MusicStreaming();
                            string[] values = xlRange.Cells[i, j].Value2.ToString().Split('\t');
                            if (values[3].Contains("10/08/2016"))
                            {
                                ms.PLAY_ID = values[0];
                                ms.CLIENT_ID = Convert.ToInt32(values[1]);
                                ms.SONG_ID = Convert.ToInt32(values[2]);
                                ms.PLAY_TS = Convert.ToDateTime(values[3]);
                                streamList.Add(ms);
                                ms = null;
                            }
                        }
                    }
                }
                
                end = DateTime.Now;
                string s = (end - start).ToString();    //time elapsed loading excel to hashset

                var clientByDistinctSong = streamList
                  .GroupBy(l => l.CLIENT_ID)
                  .Select(g => new
                  {
                      CLIENT_ID = g.Key,
                      Count = g.Select(l => l.SONG_ID).Distinct().Count()
                  });

                var desired = clientByDistinctSong
                  .GroupBy(l => l.Count)
                  .Select(g => new
                  {
                      DISTINCT_PLAY_COUNT = g.Key,
                      CLIENT_COUNT = g.Count()
                  });

                var buffer = new StringBuilder();
                buffer.AppendLine("DISTINCT_PLAY_COUNT,CLIENT_COUNT");
                desired.ToList().ForEach(item => buffer.AppendLine(String.Format("{0},{1}", item.DISTINCT_PLAY_COUNT, item.CLIENT_COUNT)));
                File.WriteAllText(@"C: \Users\***\exhibitA-output.txt", buffer.ToString());

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(excelWorksheet);
                excelWorkbook.Close();
                Marshal.ReleaseComObject(excelWorkbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        class MusicStreaming
        {
            public string PLAY_ID { get; set; }
            public int SONG_ID { get; set; }
            public int CLIENT_ID { get; set; }
            public DateTime PLAY_TS { get; set; }
        }

    }
}
