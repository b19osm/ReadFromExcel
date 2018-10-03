using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReadFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                StreamReader csvreader = new StreamReader(@"C:\Users\***\exhibitA-input.csv");
                HashSet<MusicStreaming> streamList = new HashSet<MusicStreaming>();
                string inputLine = "";               

                MusicStreaming ms = null;
                DateTime end;
                DateTime start = DateTime.Now;
                while ((inputLine = csvreader.ReadLine()) != null)
                {
                    ms = new MusicStreaming();
                    string[] values = inputLine.Split(new char[] { '\t' });
                    if (values[3].Contains("10/08/2016"))
                    {
                        ms.PLAY_ID = values[0];
                        ms.CLIENT_ID = Convert.ToInt32(values[2]);
                        ms.SONG_ID = Convert.ToInt32(values[1]);
                        ms.PLAY_TS = Convert.ToDateTime(values[3]);

                        streamList.Add(ms);
                        ms = null;
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
                desired.ToList().ForEach(item => buffer.AppendLine(String.Format("{0},\t {1}", item.DISTINCT_PLAY_COUNT, item.CLIENT_COUNT)));
                File.WriteAllText(@"C: \Users\***\exhibitA-output.txt", buffer.ToString());
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
