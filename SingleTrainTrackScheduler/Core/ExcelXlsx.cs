// ExcelXlsx.cs
// OpenXML (XLSX) için
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SingleTrainTrackScheduler.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace SingleTrainTrackScheduler.Core
{
    // Not: Sınıf CSV üretimini de içeriyor; XLSX üretimi altta.
    public static class ExcelXlsx
    {
        // --------- Ortak yardımcılar ----------
        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        private static string Escape(string v)
            => v?.Contains(',') == true ? $"\"{v.Replace("\"", "\"\"")}\"" : v ?? "";

        private static string HHMM(int m)
        {
            m = ((m % 1440) + 1440) % 1440;
            return $"{(m / 60) % 24:00}:{m % 60:00}";
        }

        private static string GetRuntimeType(Runtime r)
            => (r?.GetType().GetProperty("TType")?.GetValue(r) as string)
            ?? (r?.GetType().GetProperty("Type")?.GetValue(r) as string)
            ?? "";

        private static int GetRuntimeMinutes(Runtime r)
        {
            var p = r?.GetType().GetProperty("RunMin");
            return p != null ? (int)p.GetValue(r) : 0;
        }

        // ============================================================
        // ====================== CSV ÇIKTILARI =======================
        // ============================================================

        /// <summary>“Uzun” format CSV: train_id,type,station,arr,dep,dwell</summary>
        public static void WriteTimetableCsv(
            string path,
            IList<string> stationOrder,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, string> types,
            Dictionary<string, Dictionary<string, int>> dwells,
            IList<string> onlyStations = null,
            bool outMin = false,
            bool outHHMM = true)
        {
            var filt = onlyStations?.ToHashSet(StringComparer.OrdinalIgnoreCase);

            var sb = new StringBuilder();
            sb.AppendLine("train_id,type,station,arr,dep,dwell");

            foreach (var tid in times.Keys.OrderBy(k => k, StringComparer.Ordinal))
            {
                var ttype = types != null && types.TryGetValue(tid, out var tt) ? tt : "";
                foreach (var s in stationOrder)
                {
                    if (filt != null && !filt.Contains(s)) continue;
                    if (!times[tid].TryGetValue(s, out var a)) continue;

                    int dwell = 0;
                    if (dwells != null &&
                        dwells.TryGetValue(tid, out var map) &&
                        map.TryGetValue(s, out var d))
                        dwell = d;

                    int dep = a + dwell;

                    string fmt(int m) => outMin ? m.ToString(CultureInfo.InvariantCulture) :
                                                 outHHMM ? HHMM(m) : m.ToString(CultureInfo.InvariantCulture);

                    sb.AppendLine($"{tid},{ttype},{Escape(s)},{fmt(a)},{fmt(dep)},{dwell}");
                }
            }

            File.WriteAllText(path, sb.ToString(), Utf8NoBom);
        }

        /// <summary>Basit CSV şablon seti üret.</summary>
        public static void WriteTemplate(string path)
        {
            var dir = Path.GetDirectoryName(path) ?? ".";
            var stem = Path.GetFileNameWithoutExtension(path);
            Directory.CreateDirectory(dir);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_stations.csv"),
                new[]
                {
                    "station,km",
                    "HALKALI,0",
                    "YEŞİLKÖY,25",
                    "ATAKÖY,40",
                    "YENİMAHALLE,65",
                }, Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_segments.csv"),
                new[]
                {
                    "u,v",
                    "HALKALI,YEŞİLKÖY",
                    "YEŞİLKÖY,ATAKÖY",
                    "ATAKÖY,YENİMAHALLE",
                }, Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_runtimes.csv"),
                new[]
                {
                    "u,v,type,run_min",
                    "HALKALI,YEŞİLKÖY,P,60",
                    "YEŞİLKÖY,ATAKÖY,P,60",
                    "ATAKÖY,YENİMAHALLE,P,60",
                }, Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_trains.csv"),
                new[]
                {
                    "train_id,type,dir,origin,req_time,dest",
                    "T01,P,Up,HALKALI,60,YENİMAHALLE",
                    "T02,P,Down,YENİMAHALLE,120,HALKALI",
                }, Utf8NoBom);
        }

        /// <summary>Mevcut dataset’i CSV’lere döker (+timetable.csv,+info.txt).</summary>
        public static void WriteDataset(
            string path,
            List<Station> stations,
            List<(string u, string v)> segments,
            List<Train> trains,
            IList<Runtime> runtimes,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, string> types,
            Dictionary<string, Dictionary<string, int>> dwells,
            int headSame,
            int clearOpp,
            int stationSlack)
        {
            var dir = Path.GetDirectoryName(path) ?? ".";
            var stem = Path.GetFileNameWithoutExtension(path);
            Directory.CreateDirectory(dir);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_stations.csv"),
                new[] { "station,km" }.Concat(
                    stations.Select(s => $"{s.Name},{s.Km.ToString(CultureInfo.InvariantCulture)}")),
                Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_segments.csv"),
                new[] { "u,v" }.Concat(segments.Select(e => $"{e.u},{e.v}")), Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_trains.csv"),
                new[] { "train_id,type,dir,origin,req_time,dest" }.Concat(
                    trains.Select(t => $"{t.Id},{t.Type},{t.Direction},{t.Origin},{t.ReqTime},{t.Dest ?? ""}")),
                Utf8NoBom);

            File.WriteAllLines(Path.Combine(dir, $"{stem}_runtimes.csv"),
                new[] { "u,v,type,run_min" }.Concat(
                    (runtimes ?? Array.Empty<Runtime>()).Select(r => $"{r.U},{r.V},{GetRuntimeType(r)},{GetRuntimeMinutes(r)}")),
                Utf8NoBom);

            // timetable + info
            WriteTimetableCsv(Path.Combine(dir, $"{stem}_timetable.csv"),
                              stations.Select(s => s.Name).ToList(),
                              times, types, dwells, null, outMin: false, outHHMM: true);

            File.WriteAllText(Path.Combine(dir, $"{stem}_info.txt"),
                $"headSame={headSame}\nclearOpp={clearOpp}\nstationSlack={stationSlack}\n", Utf8NoBom);
        }

        // --- İstasyon adı normalizasyon yardımcıları ---
        private static string CanonKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var t = s.Replace('\u00A0', ' ').Trim();
            while (t.Contains("  ")) t = t.Replace("  ", " ");
            var tr = new CultureInfo("tr-TR");
            t = t.ToUpper(tr);
            return t.Normalize(NormalizationForm.FormC);
        }

        private static string FoldDiacritics(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var formD = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(formD.Length);
            foreach (var ch in formD)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        private static string ResolveStationRef(
            string raw, List<Station> stations,
            Dictionary<string, Station> byKey,
            Dictionary<string, Station> byFold)
        {
            var txt = (raw ?? "").Trim();
            if (txt.Length == 0) return txt;

            if (TryParseRowIndexLikeA1(txt, out var idx))
            {
                if (idx >= 1 && idx <= stations.Count)
                    return stations[idx - 1].Name;
            }

            var k = CanonKey(txt);
            if (byKey.TryGetValue(k, out var st1)) return st1.Name;

            var k2 = FoldDiacritics(k);
            if (byFold.TryGetValue(k2, out var st2)) return st2.Name;

            return txt;
        }

        public static (List<Station> stations,
               List<(string u, string v)> segments,
               List<Train> trains,
               List<Runtime> runtimes)
        ReadDatasetXlsx(string path)
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var wb = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart yok.");
            var sst = wb.SharedStringTablePart?.SharedStringTable;

            static int ColIndex(string cellRef)
            {
                int n = 0;
                foreach (char ch in cellRef.Where(char.IsLetter))
                    n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                return n - 1;
            }
            string CellText(Cell c)
            {
                if (c == null) return "";
                if (c.DataType != null && c.DataType == CellValues.SharedString)
                {
                    if (int.TryParse(c.CellValue?.Text, out var ix) && sst != null && ix >= 0 && ix < sst.Count())
                        return sst.ElementAt(ix).InnerText ?? "";
                    return "";
                }
                if (c.DataType != null && c.DataType == CellValues.InlineString)
                    return c.InlineString?.Text?.Text ?? "";
                return c.CellValue?.Text ?? "";
            }
            List<string[]> ReadRows(string sheetName)
            {
                var sheet = wb.Workbook.Sheets?.Elements<Sheet>()
                                .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null) return new List<string[]>();
                var wsp = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var data = wsp.Worksheet.GetFirstChild<SheetData>();
                var rows = new List<string[]>();
                foreach (var r in data.Elements<Row>())
                {
                    var list = new List<string>();
                    foreach (var c in r.Elements<Cell>())
                    {
                        var ci = ColIndex(c.CellReference!.Value!);
                        while (list.Count <= ci) list.Add("");
                        list[ci] = CellText(c);
                    }
                    rows.Add(list.ToArray());
                }
                return rows;
            }
            static int ParseTimeToMin(string s)
            {
                if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var m)) return m;
                if (!string.IsNullOrWhiteSpace(s) && s.Contains(':'))
                {
                    var p = s.Split(':');
                    if (p.Length >= 2 &&
                        int.TryParse(p[0], out var hh) &&
                        int.TryParse(p[1], out var mm))
                        return ((hh % 24 + 24) % 24) * 60 + ((mm % 60 + 60) % 60);
                }
                return 0;
            }

            // ---- stations ----
            var stRows = ReadRows("stations");
            var stations = stRows
                .Skip(1)
                .Where(r => r.Length >= 2 && !string.IsNullOrWhiteSpace(r[0]))
                .Select(r =>
                {
                    var name = (r[0] ?? "").Replace('\u00A0', ' ').Trim();
                    while (name.Contains("  ")) name = name.Replace("  ", " ");
                    var kmOk = double.TryParse(r[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var km);
                    return new Station(name, kmOk ? km : 0);
                })
                .OrderBy(s => s.Km)
                .ToList();

            var byKey = new Dictionary<string, Station>(StringComparer.Ordinal);
            var byFold = new Dictionary<string, Station>(StringComparer.Ordinal);
            foreach (var s in stations)
            {
                var k = CanonKey(s.Name);
                var k2 = FoldDiacritics(k);
                if (!byKey.ContainsKey(k)) byKey[k] = s;
                if (!byFold.ContainsKey(k2)) byFold[k2] = s;
            }

            // ---- segments ----
            var segRows = ReadRows("segments");
            var segments = segRows.Skip(1)
                .Where(r => r.Length >= 2 && (!string.IsNullOrWhiteSpace(r[0]) || !string.IsNullOrWhiteSpace(r[1])))
                .Select(r =>
                {
                    var u = ResolveStationRef(r[0], stations, byKey, byFold);
                    var v = ResolveStationRef(r[1], stations, byKey, byFold);
                    return (u, v);
                })
                .Where(t => !string.IsNullOrWhiteSpace(t.u) && !string.IsNullOrWhiteSpace(t.v))
                .ToList();

            // ---- trains ----  (BAŞLIKLARA GÖRE GÜNCEL OKUYUCU)
            var trRows = ReadRows("trains");

            int Idx(string[] hdr, params string[] names)
            {
                for (int i = 0; i < hdr.Length; i++)
                    foreach (var n in names)
                        if (string.Equals(hdr[i]?.Trim(), n, StringComparison.OrdinalIgnoreCase))
                            return i;
                return -1;
            }

            var hdrTr = trRows.FirstOrDefault() ?? Array.Empty<string>();
            int ixId = Idx(hdrTr, "train_id", "id");
            int ixType = Idx(hdrTr, "type");
            int ixDir = Idx(hdrTr, "dir", "direction");
            int ixOrg = Idx(hdrTr, "origin", "org");
            int ixReq = Idx(hdrTr, "req_time", "req", "time");
            int ixDest = Idx(hdrTr, "dest", "destination");

            var trains = trRows.Skip(1)
                .Where(r => r.Length > 0 && (ixId >= 0 && ixId < r.Length) && !string.IsNullOrWhiteSpace(r[ixId]))
                .Select(r =>
                {
                    string id = (r[ixId] ?? "").Trim();
                    string typ = (ixType >= 0 && ixType < r.Length) ? (r[ixType] ?? "").Trim() : "";
                    string dir = (ixDir >= 0 && ixDir < r.Length) ? (r[ixDir] ?? "").Trim() : "";
                    string orgRaw = (ixOrg >= 0 && ixOrg < r.Length) ? r[ixOrg] : "";
                    string dstRaw = (ixDest >= 0 && ixDest < r.Length) ? r[ixDest] : "";
                    string origin = ResolveStationRef(orgRaw, stations, byKey, byFold);
                    string dest = ResolveStationRef(dstRaw, stations, byKey, byFold);
                    int req = ParseTimeToMin((ixReq >= 0 && ixReq < r.Length) ? (r[ixReq] ?? "0") : "0");

                    if (string.IsNullOrWhiteSpace(dest))
                        dest = (dir ?? "").Trim().Equals("up", StringComparison.OrdinalIgnoreCase)
                             ? stations.Last().Name
                             : stations.First().Name;

                    return new Train(id, typ, dir, origin, req, dest);
                })
                .ToList();

            // ---- runtimes ----
            var rtRows = ReadRows("runtimes");
            var runtimes = rtRows.Skip(1)
                .Where(r => r.Length >= 4 && (!string.IsNullOrWhiteSpace(r[0]) || !string.IsNullOrWhiteSpace(r[1])))
                .Select(r =>
                {
                    var u = ResolveStationRef(r[0], stations, byKey, byFold);
                    var v = ResolveStationRef(r[1], stations, byKey, byFold);
                    var type = (r[2] ?? "").Trim();
                    var run = int.TryParse(r[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var m) ? m : 0;
                    return new Runtime(u, v, type, run);
                })
                .Where(rt => !string.IsNullOrWhiteSpace(rt.U) && !string.IsNullOrWhiteSpace(rt.V))
                .ToList();

            // ---- stations listesine veri sayfalarında geçen ama eksik olan adları ekle ----
            var usedNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (u, v) in segments) { usedNames.Add(u); usedNames.Add(v); }
            foreach (var t in trains)
            {
                if (!string.IsNullOrWhiteSpace(t.Origin)) usedNames.Add(t.Origin);
                if (!string.IsNullOrWhiteSpace(t.Dest)) usedNames.Add(t.Dest);
            }
            foreach (var r in runtimes) { usedNames.Add(r.U); usedNames.Add(r.V); }

            var present = new HashSet<string>(stations.Select(s => CanonKey(s.Name)), StringComparer.Ordinal);
            double nextKm = stations.Count > 0 ? stations.Max(s => s.Km) + 1 : 0;

            foreach (var nm in usedNames)
            {
                var k = CanonKey(nm);
                if (!present.Contains(k))
                {
                    var sNew = new Station(nm, nextKm);
                    stations.Add(sNew);
                    present.Add(k);

                    var kk = CanonKey(sNew.Name);
                    var kk2 = FoldDiacritics(kk);
                    if (!byKey.ContainsKey(kk)) byKey[kk] = sNew;
                    if (!byFold.ContainsKey(kk2)) byFold[kk2] = sNew;

                    nextKm += 1.0;
                }
            }

            stations = stations.OrderBy(s => s.Km).ToList();

            return (stations, segments, trains, runtimes);
        }

        // "12", "F12", "$F$12", "AB7" gibi A1 adreslerinden sonundaki satır numarasını al
        private static bool TryParseRowIndexLikeA1(string s, out int idx)
        {
            idx = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;

            var t = s.Replace("$", "").Trim();

            int i = t.Length - 1;
            while (i >= 0 && char.IsDigit(t[i])) i--;
            int start = i + 1;
            if (start >= t.Length) return false;
            var digits = t.Substring(start);
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx);
        }

        /// <summary>WriteTemplate/WriteDataset ile üretilen CSV setini geri oku.</summary>
        public static (List<Station> stations,
                       List<(string u, string v)> segments,
                       List<Train> trains,
                       List<Runtime> runtimes)
        ReadDataset(string path)
        {
            var dir = Path.GetDirectoryName(path) ?? ".";
            var stem = Path.GetFileNameWithoutExtension(path);

            var stations = ExcelIO.ReadStations(Path.Combine(dir, $"{stem}_stations.csv"));
            var segments = ExcelIO.ReadSegments(Path.Combine(dir, $"{stem}_segments.csv"));
            var trains = ExcelIO.ReadTrains(Path.Combine(dir, $"{stem}_trains.csv"));
            var runtimes = ExcelIO.ReadRuntimes(Path.Combine(dir, $"{stem}_runtimes.csv"));

            return (stations, segments, trains, runtimes);
        }

        // ============================================================
        // ====================== XLSX ÇIKTILARI ======================
        // ============================================================

        private static string ColName(int index0)
        {
            int n = index0 + 1;
            string s = "";
            while (n > 0)
            {
                n--;
                s = (char)('A' + (n % 26)) + s;
                n /= 26;
            }
            return s;
        }

        private static void AddSheet(WorkbookPart wb, ref uint sid, string name, IEnumerable<string[]> rows)
        {
            var wsp = wb.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());
            wsp.Worksheet.Save();

            if (wb.Workbook.Sheets == null)
                wb.Workbook.AppendChild(new Sheets());

            var sheets = wb.Workbook.Sheets;
            var relId = wb.GetIdOfPart(wsp);

            sheets.Append(new Sheet
            {
                Id = relId,
                SheetId = sid++,
                Name = name
            });

            var data = wsp.Worksheet.GetFirstChild<SheetData>();

            uint r = 1;
            foreach (var arr in rows)
            {
                var row = new Row { RowIndex = r };
                for (int c = 0; c < arr.Length; c++)
                {
                    var cell = new Cell
                    {
                        CellReference = ColName(c) + r.ToString(CultureInfo.InvariantCulture),
                        DataType = CellValues.String,
                        CellValue = new CellValue(arr[c] ?? "")
                    };
                    row.Append(cell);
                }
                data.Append(row);
                r++;
            }

            wsp.Worksheet.Save();
        }

        private static IEnumerable<string[]> BuildTimetableLongRows(
            IList<Station> stations,
            IEnumerable<Train> trains,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, Dictionary<string, int>> dwells,
            bool outMin = false, bool outHHMM = true)
        {
            yield return new[] { "train_id", "type", "station", "arr", "dep", "dwell" };

            string Fmt(int m) => outMin ? m.ToString(CultureInfo.InvariantCulture)
                                        : outHHMM ? HHMM(m)
                                                  : m.ToString(CultureInfo.InvariantCulture);

            foreach (var t in trains.OrderBy(z => z.ReqTime).ThenBy(z => z.Id, StringComparer.Ordinal))
            {
                foreach (var st in stations)
                {
                    if (!times.TryGetValue(t.Id, out var map) || !map.TryGetValue(st.Name, out var a))
                        continue;

                    int dwell = 0;
                    if (dwells != null &&
                        dwells.TryGetValue(t.Id, out var dmap) &&
                        dmap.TryGetValue(st.Name, out var d))
                        dwell = d;

                    int dep = a + dwell;
                    yield return new[] { t.Id, t.Type, st.Name, Fmt(a), Fmt(dep), dwell.ToString(CultureInfo.InvariantCulture) };
                }
            }
        }

        private static IEnumerable<string[]> BuildTimetableWideRows(
            IList<Station> stations,
            IEnumerable<Train> trains,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, Dictionary<string, int>> dwells,
            bool outMin = false, bool outHHMM = true)
        {
            var header = new List<string> { "train_id", "type", "dir", "origin", "req_time", "dest" };
            foreach (var s in stations)
            {
                header.Add($"{s.Name} (arr)");
                header.Add($"{s.Name} (dep)");
                header.Add($"{s.Name} (dwell)");
            }
            yield return header.ToArray();

            string Fmt(int m) => outMin ? m.ToString(CultureInfo.InvariantCulture)
                                        : outHHMM ? HHMM(m)
                                                  : m.ToString(CultureInfo.InvariantCulture);

            foreach (var t in trains.OrderBy(z => z.ReqTime).ThenBy(z => z.Id, StringComparer.Ordinal))
            {
                var row = new List<string> { t.Id, t.Type, t.Direction, t.Origin, Fmt(t.ReqTime), t.Dest ?? "" };

                foreach (var st in stations)
                {
                    if (times.TryGetValue(t.Id, out var map) && map.TryGetValue(st.Name, out var arr))
                    {
                        int dwell = 0;
                        if (dwells != null &&
                            dwells.TryGetValue(t.Id, out var dmap) &&
                            dmap.TryGetValue(st.Name, out var d))
                            dwell = d;

                        int dep = arr + dwell;
                        row.Add(Fmt(arr));
                        row.Add(Fmt(dep));
                        row.Add(dwell.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        row.Add(""); row.Add(""); row.Add("");
                    }
                }

                yield return row.ToArray();
            }
        }

        private static void AddTimetableWideSheet(
            WorkbookPart wb, ref uint sid,
            List<Station> stations,
            List<Train> trains,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, Dictionary<string, int>> dwells)
        {
            var rows = BuildTimetableWideRows(stations, trains, times, dwells, outMin: false, outHHMM: true);
            AddSheet(wb, ref sid, "timetable_wide", rows);
        }

        /// <summary>
        /// Tüm veriyi tek bir .xlsx’e yazar:
        /// stations, segments, runtimes, trains, timetable, timetable_wide, info
        /// </summary>
        public static void WriteDatasetXlsx(
            string path,
            List<Station> stations,
            List<(string u, string v)> segments,
            List<Train> trains,
            IList<Runtime> runtimes,
            Dictionary<string, Dictionary<string, int>> times,
            Dictionary<string, string> types,
            Dictionary<string, Dictionary<string, int>> dwells,
            int headSame,
            int clearOpp,
            int stationSlack)
        {
            using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            var wb = doc.AddWorkbookPart();
            wb.Workbook = new Workbook();
            wb.Workbook.AppendChild(new Sheets());

            uint sid = 1;

            // stations
            var rowsStations = new List<string[]> { new[] { "station", "km" } };
            rowsStations.AddRange(stations.Select(s => new[] { s.Name, s.Km.ToString(CultureInfo.InvariantCulture) }));
            AddSheet(wb, ref sid, "stations", rowsStations);

            // segments
            var rowsSeg = new List<string[]> { new[] { "u", "v" } };
            rowsSeg.AddRange(segments.Select(e => new[] { e.u, e.v }));
            AddSheet(wb, ref sid, "segments", rowsSeg);

            // runtimes
            var rowsRt = new List<string[]> { new[] { "u", "v", "type", "run_min" } };
            rowsRt.AddRange((runtimes ?? Array.Empty<Runtime>())
                .Select(r => new[] { r.U, r.V, GetRuntimeType(r), GetRuntimeMinutes(r).ToString(CultureInfo.InvariantCulture) }));
            AddSheet(wb, ref sid, "runtimes", rowsRt);

            // trains
            var rowsTr = new List<string[]> { new[] { "train_id", "type", "dir", "origin", "req_time", "dest" } };
            rowsTr.AddRange(trains.Select(t => new[]
            {
                t.Id, t.Type, t.Direction, t.Origin,
                t.ReqTime.ToString(CultureInfo.InvariantCulture), t.Dest ?? ""
            }));
            AddSheet(wb, ref sid, "trains", rowsTr);

            // timetable (long)
            var longRows = BuildTimetableLongRows(stations, trains, times, dwells, outMin: false, outHHMM: true);
            AddSheet(wb, ref sid, "timetable", longRows);

            // timetable_wide (Arr/Dep/Dwell)
            AddTimetableWideSheet(wb, ref sid, stations, trains, times, dwells);

            // info
            var rowsInfo = new List<string[]>
            {
                new[] { "key", "value" },
                new[] { "headSame", headSame.ToString(CultureInfo.InvariantCulture) },
                new[] { "clearOpp", clearOpp.ToString(CultureInfo.InvariantCulture) },
                new[] { "stationSlack", stationSlack.ToString(CultureInfo.InvariantCulture) },
                new[] { "generated_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
            };
            AddSheet(wb, ref sid, "info", rowsInfo);

            wb.Workbook.Save();
        }
    }
}
