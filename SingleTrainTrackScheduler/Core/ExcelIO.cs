using SingleTrainTrackScheduler.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace SingleTrainTrackScheduler.Core
{
    public static class ExcelIO
    {
        // CSV’yi önce UTF-8 (BOM’suz), olmazsa Windows-1254 (Türkçe) olarak oku
        private static string ReadTextWithTrFallback(string path)
        {
            var bytes = File.ReadAllBytes(path);
            var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            string s = utf8.GetString(bytes);

            // Replacement char çoksa 1254’e düş
            if (s.Contains('\uFFFD'))
            {
                s = Encoding.GetEncoding(1254).GetString(bytes);
            }
            return s;
        }

        private static IEnumerable<string> ReadCsvLines(string path)
        {
            string text = ReadTextWithTrFallback(path);
            return text.Replace("\r\n", "\n").Replace("\r", "\n")
                       .Split('\n').Select(l => l.TrimEnd());
        }

        // Data\stations.csv  — header: station,km
        public static List<Station> ReadStations(string path)
        {
            var list = new List<Station>();
            foreach (var line in ReadCsvLines(path).Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var c = line.Split(',');
                string name = (c[0] ?? "").Trim();
                double km = double.Parse(c[1], CultureInfo.InvariantCulture);
                list.Add(new Station(name, km));
            }
            return list.OrderBy(s => s.Km).ToList();
        }

        // Data\segments.csv — header: u,v
        public static List<(string u, string v)> ReadSegments(string path)
        {
            var list = new List<(string, string)>();
            foreach (var line in ReadCsvLines(path).Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var c = line.Split(',');
                list.Add((c[0].Trim(), c[1].Trim()));
            }
            return list;
        }

        // Data\runtimes.csv — header: u,v,type,run_min
        public static List<Runtime> ReadRuntimes(string path)
        {
            var list = new List<Runtime>();
            foreach (var line in ReadCsvLines(path).Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var c = line.Split(',');
                list.Add(new Runtime(c[0].Trim(), c[1].Trim(), c[2].Trim(),
                                     int.Parse(c[3], CultureInfo.InvariantCulture)));
            }
            return list;
        }

        // Data\trains.csv — header: train_id,type,dir,origin,req_time,dest
        public static List<Train> ReadTrains(string path)
        {
            var list = new List<Train>();
            foreach (var line in ReadCsvLines(path).Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var c = line.Split(',');
                var dest = c.Length >= 6 ? c[5].Trim() : "";
                list.Add(new Train(
                    c[0].Trim(), c[1].Trim(), c[2].Trim(), c[3].Trim(),
                    int.Parse(c[4], CultureInfo.InvariantCulture), dest));
            }
            return list;
        }
    }
}
