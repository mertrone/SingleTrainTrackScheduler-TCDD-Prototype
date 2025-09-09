using Microsoft.Win32;
using OxyPlot;
using SingleTrainTrackScheduler.Core;
using SingleTrainTrackScheduler.Interop;   // <-- NativeOrtools.EnsureLoaded() için
using SingleTrainTrackScheduler.Models;
using SingleTrainTrackScheduler.Optimizer;
using SingleTrainTrackScheduler.Rendering;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PdfExporterManaged = OxyPlot.PdfExporter;       // Skia ile çakışmasın diye alias


namespace SingleTrainTrackScheduler
{
    public partial class MainWindow : Window
    {
        private bool _warnedMissingRt = false;

        // ——— Mor uyarı penceresi
        private WarningsWindow _warnWin;

        // kaydırma eşiği (P/YHT=15, Yük=45)
        private static int ShiftLimit(Train t)
        {
            var tt = (t.Type ?? "").Trim().ToUpperInvariant();
            if (tt == "F" || tt == "G") return 45;  // yük
            return 15;                               // yolcu + YHT
        }

        // başlangıç referans saatleri (mor kontrol için)
        private readonly Dictionary<string, int> _baseReq = new();

        private static readonly OxyColor Violet = OxyColor.FromRgb(128, 0, 128);

        // sn olarak 0.00 biçiminde süre yazdır
        private static string Sec(Stopwatch sw) => $"{sw.Elapsed.TotalSeconds:0.00} sn";

        // sınıf alanları arasına ekle
        private const int MAX_DWELL_PER_STATION = 20;

        // ÖNERİ maliyet katsayıları: dwell kaydırmadan DAHA tercih edilir
        private const int COST_SHIFT = 10; // Kaydırma pahalı
        private const int COST_DWELL = 1;  // Dwell ucuz (öneride önce dwell)

        private Engine _engine;
        private PlotRenderer _plot;

        // durum
        private readonly HashSet<string> _unscheduled = new();
        private List<Conflict> _lastConflicts = new();

        // son çizimde limit aşımı uyarı metinleri
        private List<string> _lastShiftWarnings = new();

        public MainWindow()
        { //NativeOrtools.EnsureLoaded();
            InitializeComponent();
        }

        private string DataPath(string file) =>
            System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", file);

        // Tren path'ini, Times'ta zamanı olan ilk kesite kadar kırpar
        private List<string> TrimToComputed(string tid)
        {
            var names = _engine.Paths[tid];
            var computed = _engine.Times.TryGetValue(tid, out var map) ? map : null;
            var list = new List<string>();
            if (computed == null) return list;

            foreach (var n in names)
            {
                if (!computed.ContainsKey(n)) break;   // buradan sonrası hesaplanmamış → kes
                list.Add(n);
            }
            // en az iki istasyon olmalı ki bir segment çizilebilsin
            if (list.Count < 2) list.Clear();
            return list;
        }


        // Önerileri ekranda kısa bir özetle göster (tam liste dosyada)
        private void ShowSuggestionsInline(IEnumerable<string> lines, string filePath, string title)
        {
            var arr = lines?.ToList() ?? new List<string>();
            int preview = Math.Min(15, arr.Count); // ekranda en fazla 15 satır
            string msg = (preview > 0)
                ? string.Join(Environment.NewLine, arr.Take(preview))
                : "Öneri bulunamadı.";

            if (arr.Count > preview)
                msg += $"\n... (+{arr.Count - preview} öneri daha)";

            if (!string.IsNullOrWhiteSpace(filePath))
                msg += $"\n\nTam liste: {filePath}";

            MessageBox.Show(msg, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // {u,v} segmentine bu tren hangi uçtan giriyor? (o istasyon adı)
        private string EntryStationFor(string tid, (string u, string v) seg)
        {
            var path = _engine.Paths[tid];
            int i = path.IndexOf(seg.u);
            if (i >= 0 && i + 1 < path.Count && path[i + 1] == seg.v) return seg.u;
            i = path.IndexOf(seg.v);
            if (i >= 0 && i + 1 < path.Count && path[i + 1] == seg.u) return seg.v;
            return null; // olağan dışı: tren bu segmentte yok
        }

        // Bu trenin bir istasyondaki kullanılabilir dwell boşluğu (üst sınır 20 ise: 20 - mevcut)
        private int DwellHeadroom(string tid, string station)
        {
            if (string.IsNullOrEmpty(station)) return 0;
            var t = _engine.Trains[tid];
            int cur = t.Dwell.TryGetValue(station, out var d) ? d : 0;
            int room = MAX_DWELL_PER_STATION - cur;
            return room < 0 ? 0 : room;
        }

        // Bir çatışmada ÖNCE kimin gecikmesi gerektiğini ve "ihtiyaç" dakikayı hesapla (mevcut mantığın aynısı)
        private (string delayTid, int need) DelayTargetAndNeed(Conflict c, int headSame, int clearOpp)
        {
            // s/e pencerelerini bul
            var phys = PhysKey(c.Seg.u, c.Seg.v);

            (int s, int e) Win(string tid)
            {
                var w = _engine.SegmentWindows(tid).FirstOrDefault(z => PhysKey(z.seg.u, z.seg.v) == phys);
                if (w.seg.u != null) return (w.s, w.e);

                // emniyetli yedek
                return (c.Interval.s, c.Interval.e);
            }

            var (sA, eA) = Win(c.A);
            var (sB, eB) = Win(c.B);

            int prA = TrainPriority(_engine.Trains[c.A]);
            int prB = TrainPriority(_engine.Trains[c.B]);

            bool same = c.SameDir;
            int need = 0;
            string delay = c.B;

            if (same)
            {
                bool aLead = sA <= sB;
                int leadIn = aLead ? sA : sB;
                int leadOut = aLead ? eA : eB;
                int folIn = aLead ? sB : sA;
                int folOut = aLead ? eB : eA;

                int d1 = headSame - (folIn - leadIn);
                int d2 = headSame - (folOut - leadOut);
                need = Math.Max(0, Math.Max(d1, d2));

                if (prA > prB) delay = c.A;              // düşük öncelikli geciksin
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }
            else
            {
                // karşı yön
                if (eA <= sB) need = Math.Max(0, clearOpp - (sB - eA));
                else need = Math.Max(0, clearOpp - (sA - eB));

                if (prA > prB) delay = c.A;
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }

            if (need <= 0) need = 1; // ilerleme garantisi
            return (delay, need);
        }

        // Bir çatışma için "duruş mu / kaydırma mı / hibrit mi" kararını üret — maliyet tabanlı
        private (string text, string tid, string station, int addDwell, int addShift)
            PickBestFix(Conflict c, int headSame, int clearOpp)
        {
            var (delayTid, need) = DelayTargetAndNeed(c, headSame, clearOpp);
            var entry = EntryStationFor(delayTid, c.Seg);
            int room = DwellHeadroom(delayTid, entry);

            // Tamamı kaydırma maliyeti
            int costAllShift = need * COST_SHIFT;

            if (room >= need)
            {
                // Tümü dwell olabilir; dwell ucuz olduğundan dwell'i tercih et
                int costAllDwell = need * COST_DWELL;
                if (costAllDwell <= costAllShift && !string.IsNullOrEmpty(entry))
                {
                    return ($"• {c.Seg.u}-{c.Seg.v}: {c.A} ↔ {c.B} — {delayTid} için {entry}’de **{need} dk duruş** ekleyin.",
                            delayTid, entry, need, 0);
                }
                else
                {
                    return ($"• {c.Seg.u}-{c.Seg.v}: {c.A} ↔ {c.B} — {delayTid} başlangıcını **{need} dk kaydırın**.",
                            delayTid, entry, 0, need);
                }
            }
            else if (room > 0)
            {
                // Hibrit: (room dwell + shift)
                int shift = need - room;
                int costHybrid = room * COST_DWELL + shift * COST_SHIFT;

                if (!string.IsNullOrEmpty(entry) && costHybrid <= costAllShift)
                {
                    return ($"• {c.Seg.u}-{c.Seg.v}: {c.A} ↔ {c.B} — {delayTid} için **{entry}’de {room} dk duruş + {shift} dk kaydırma** önerilir.",
                            delayTid, entry, room, shift);
                }
                else
                {
                    return ($"• {c.Seg.u}-{c.Seg.v}: {c.A} ↔ {c.B} — {delayTid} başlangıcını **{need} dk kaydırın**.",
                            delayTid, entry, 0, need);
                }
            }
            else
            {
                // Duruş boşluğu yok → kaydır
                return ($"• {c.Seg.u}-{c.Seg.v}: {c.A} ↔ {c.B} — {delayTid} başlangıcını **{need} dk kaydırın** (duruş boşluğu yok).",
                        delayTid, entry, 0, need);
            }
        }

        // CSV'den gelen tüm trenler için: her istasyondaki toplam duruşu [0..MAX_DWELL_PER_STATION] aralığına sık.
        private void ClampAllDwellsToCap()
        {
            foreach (var tr in _engine.Trains.Values)
            {
                // trenin geçtiği istasyonlarda dolaş
                foreach (var s in _engine.Paths[tr.Id])
                {
                    if (!tr.Dwell.TryGetValue(s, out var d)) continue; // duruş tanımlı değilse geç
                    if (d < 0) d = 0;
                    if (d > MAX_DWELL_PER_STATION) d = MAX_DWELL_PER_STATION;
                    tr.Dwell[s] = d;
                }
            }
            _engine.ComputeTimes(); // zinciri ileri güncelle
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // grafik
            _plot = new PlotRenderer(_engine?.Stations?.Select(s => (s.Name, s.Km)) ?? new List<(string, double)> { ("", 0) });
            plotView.Model = _plot.Model;
            _plot.AttachTo(plotView);

            // grafikten tren seçilince solda da seç
            _plot.OnSeriesClicked = tid =>
            {
                var orderedIds = _engine.Trains.Keys.OrderBy(k => k).ToList();
                var ix = orderedIds.IndexOf(tid);
                if (ix >= 0) lstTrains.SelectedIndex = ix;
                _plot.SetSelectedTrain(tid);
            };

            LoadCsvDataAndRender();
        }

        private void LstTrains_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_engine == null) return;

            if (lstTrains.SelectedIndex >= 0)
            {
                var tid = _engine.Trains.Keys.OrderBy(k => k).ElementAt(lstTrains.SelectedIndex);
                _plot.SetSelectedTrain(tid);
            }
            else
            {
                _plot.SetSelectedTrain(null);
            }
        }

        private void LoadCsvDataAndRender()
        {
            try
            {
                var stations = ExcelIO.ReadStations(DataPath("stations.csv"));
                var segments = ExcelIO.ReadSegments(DataPath("segments.csv"));
                var runtimes = ExcelIO.ReadRuntimes(DataPath("runtimes.csv"));
                var trains = ExcelIO.ReadTrains(DataPath("trains.csv"));

                _engine = new Engine(stations, segments, trains, runtimes);
                _engine.ComputeTimes();

                // referans başlangıç saatlerini kaydet (mor için)
                _baseReq.Clear();
                foreach (var kv in _engine.Trains)
                    _baseReq[kv.Key] = kv.Value.ReqTime;

                ClampAllDwellsToCap(); // CSV'den gelen tüm trenler için duruşları sıkıştır

                // sol panel ve combobox
                lstTrains.Items.Clear();
                foreach (var t in _engine.Trains.Values.OrderBy(t => t.ReqTime))
                    lstTrains.Items.Add(
                    $"{t.Id} [{t.Type},{t.Direction}] {t.Origin}->" +
                    $"{(string.IsNullOrWhiteSpace(t.Dest) ? _engine.Paths[t.Id].Last() : t.Dest)} @{HHMM(t.ReqTime)}");

                cbStation.Items.Clear();
                foreach (var s in _engine.Stations) cbStation.Items.Add(s.Name);
                if (cbStation.Items.Count > 0) cbStation.SelectedIndex = 0;

                // grafik
                _plot.ResetStations(_engine.Stations.Select(s => (s.Name, s.Km)));
                _plot.ResetView12h();
                RefreshPlot();
                txtConfMain.Text = "Dataset yüklendi.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Yükleme Hatası", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            _warnedMissingRt = false;   // LoadCsvDataAndRender ve LoadExcelDataset sonunda

        }

        // ==== ÇİZİM ====
        // tip önceliği: 0=YHT, 1=Yolcu (P), 2=Yük (F/G)
        private static int TrainPriority(Train t)
        {
            var tt = (t.Type ?? "").Trim().ToUpperInvariant();
            if (tt is "YHT" or "Y" or "HSR" or "H") return 0;
            if (tt is "P" or "YOLCU") return 1;
            return 2; // F, G
        }

        // yönsüz segment anahtarı
        private static (string a, string b) PhysKey(string u, string v)
            => string.CompareOrdinal(u, v) < 0 ? (u, v) : (v, u);

        // ---- kaydırma limitleri (mor çizim için) ----
        private static int MaxShiftLimitForType(string type)
        {
            var tt = (type ?? "").Trim().ToUpperInvariant();
            // Yük (F/G) = 45 dk, Yolcu & YHT = 15 dk
            if (tt == "F" || tt == "G") return 45;
            return 15;
        }

        private bool ExceedsShiftLimit(string tid, out int deltaAbs, out int limit)
        {
            deltaAbs = 0; limit = 0;
            if (!_baseReq.TryGetValue(tid, out var baseReq)) return false;
            var t = _engine.Trains[tid];
            int diff = t.ReqTime - baseReq;
            deltaAbs = Math.Abs(diff);
            limit = MaxShiftLimitForType(t.Type);
            return deltaAbs > limit;
        }

        // Sırayla yerleştir: YHT→Yolcu→Yük ve erken saate göre.
        // Yolcu/YHT çatışırsa çatışma başından itibaren kesikli; Yük çatışırsa atlanır.
        private (List<string> placed, Dictionary<string, int> dashFrom, List<string> skipped)
        BuildPlacementPlan(int headSame, int clearOpp, int stationSlack)
        {
            var order = _engine.Trains.Values
                .OrderBy(t => TrainPriority(t))
                .ThenBy(t => t.ReqTime)
                .Select(t => t.Id).ToList();

            var allConfs = _engine.DetectConflicts(headSame, clearOpp, stationSlack);
            var placed = new List<string>();
            var dashFrom = new Dictionary<string, int>();
            var skipped = new List<string>();

            foreach (var tid in order)
            {
                var myConfs = allConfs.Where(c => (c.A == tid && placed.Contains(c.B)) ||
                                                  (c.B == tid && placed.Contains(c.A)))
                                      .ToList();

                if (myConfs.Count == 0) { placed.Add(tid); continue; }

                var prio = TrainPriority(_engine.Trains[tid]);
                if (prio >= 2)
                {
                    // Yük treni çatışıyorsa görünür olsun: tam hat KESİKLİ ve KIRMIZI çizdireceğiz
                    dashFrom[tid] = myConfs.Min(c => c.Interval.s);
                    placed.Add(tid);
                    continue;
                }

                // Yolcu/YHT → kesikli: çatışmanın başladığı en erken an
                dashFrom[tid] = myConfs.Min(c => c.Interval.s);
                placed.Add(tid);
            }

            return (placed, dashFrom, skipped);
        }

        // Çatışmalara göre öneri üret (dakika cinsinden)
        private List<string> MakeSuggestions(int headSame, int clearOpp, int stationSlack)
        {
            var lines = new List<string>();
            var confs = _engine.DetectConflicts(headSame, clearOpp, stationSlack)
                               .OrderBy(c => c.Interval.s).ToList();

            foreach (var c in confs)
            {
                var fix = PickBestFix(c, headSame, clearOpp);
                lines.Add(fix.text);
            }

            if (lines.Count == 0) lines.Add("Çatışma yok; öneri gerekmez.");
            return lines;
        }

        // Bu fiziksel segmente girerken bekleme yapılacak istasyonu bul
        private string EnterStationForSegment(string tid, (string a, string b) phys)
        {
            var path = _engine.Paths[tid];
            int ix = path.IndexOf(phys.a);
            if (ix >= 0 && ix + 1 < path.Count && path[ix + 1] == phys.b) return phys.a;
            ix = path.IndexOf(phys.b);
            if (ix >= 0 && ix + 1 < path.Count && path[ix + 1] == phys.a) return phys.b;
            return phys.a; // emniyetli varsayılan
        }

        // Bir çatışmada HANGİ tren ne kadar gecikmeli? (öneriyle aynı mantık)
        private (string delayTid, int need) ComputeDelayForConflict(
            Conflict c, int headSame, int clearOpp, int stationSlack)
        {
            var phys = PhysKey(c.Seg.u, c.Seg.v);

            (int s, int e) Win(string tid)
            {
                var win = _engine.SegmentWindows(tid)
                                 .FirstOrDefault(w => PhysKey(w.seg.u, w.seg.v) == phys);
                if (!(win.seg.u is null) && !(win.seg.v is null)) return (win.s, win.e);

                var path = _engine.Paths[tid];
                int idx = path.IndexOf(phys.a);
                if (idx >= 0 && idx + 1 < path.Count && path[idx + 1] == phys.b)
                {
                    int aA = _engine.Times[tid][phys.a];
                    int dwA = _engine.Trains[tid].Dwell.TryGetValue(phys.a, out var dA) ? dA : 0;
                    int dA2 = aA + dwA;
                    int aB = _engine.Times[tid][phys.b];
                    return (dA2, aB);
                }
                idx = path.IndexOf(phys.b);
                if (idx >= 0 && idx + 1 < path.Count && path[idx + 1] == phys.a)
                {
                    int aB = _engine.Times[tid][phys.b];
                    int dwB = _engine.Trains[tid].Dwell.TryGetValue(phys.b, out var dB) ? dB : 0;
                    int dB2 = aB + dwB;
                    int aA = _engine.Times[tid][phys.a];
                    return (dB2, aA);
                }
                return (c.Interval.s, c.Interval.e);
            }

            var (sA, eA) = Win(c.A);
            var (sB, eB) = Win(c.B);

            int prA = TrainPriority(_engine.Trains[c.A]);
            int prB = TrainPriority(_engine.Trains[c.B]);

            bool same = c.SameDir;
            int need = 0;
            string delay = c.B;

            if (same)
            {
                bool aLead = sA <= sB;
                int leadIn = aLead ? sA : sB;
                int leadOut = aLead ? eA : eB;
                int folIn = aLead ? sB : sA;
                int folOut = aLead ? eB : eA;

                int d1 = headSame - (folIn - leadIn);
                int d2 = headSame - (folOut - leadOut);
                need = Math.Max(0, Math.Max(d1, d2));

                if (prA > prB) delay = c.A;
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }
            else
            {
                if (eA <= sB) need = Math.Max(0, clearOpp - (sB - eA));
                else need = Math.Max(0, clearOpp - (sA - eB));

                if (prA > prB) delay = c.A;
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }

            if (need <= 0) need = 1; // ilerleme garantisi
            return (delay, need);
        }

        private string WriteSuggestionsToDesktop(IEnumerable<string> lines)
        {
            var path = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                $"optimizer_suggestions_{DateTime.Now:yyyyMMdd_HHmm}.txt");
            System.IO.File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
            return path;
        }

        private static OxyColor ColorByType(string ttype)
        {
            ttype = (ttype ?? "P").ToUpperInvariant();
            return ttype switch
            {
                "F" => OxyColors.Black,
                "G" => OxyColor.FromRgb(0x33, 0x55, 0xE6),
                _ => OxyColor.FromRgb(0x18, 0x9A, 0x2B)
            };
        }

        private (List<Conflict> confs, Dictionary<string, List<((string u, string v) seg, int s, int e)>> overlays)
            BuildOverlays(int head, int clear, int slack)
        {
            var confs = _engine.DetectConflicts(head, clear, slack);
            var overlays = new Dictionary<string, List<((string, string), int, int)>>();

            foreach (var c in confs)
            {
                if (!overlays.ContainsKey(c.A))
                    overlays[c.A] = new List<((string, string), int, int)>();
                if (!overlays.ContainsKey(c.B))
                    overlays[c.B] = new List<((string, string), int, int)>();

                overlays[c.A].Add((c.Seg, c.Interval.s, c.Interval.e));
                overlays[c.B].Add((c.Seg, c.Interval.s, c.Interval.e));
            }

            return (confs, overlays);
        }

        private void RefreshPlot()
        {
            _engine.ComputeTimes();
            if (!_warnedMissingRt && _engine.LastMissingRuntimes.Count > 0)
            {
                var gr = _engine.LastMissingRuntimes
                    .GroupBy(m => $"{m.Type} {m.U}-{m.V}")
                    .Select(g => $"{g.Key}  (adet: {g.Count()})")
                    .OrderBy(s => s)
                    .ToList();

                MessageBox.Show(
                    "Runtimes sayfasında aşağıdaki (u,v,type) çiftleri eksik.\n" +
                    "Bu segmentlerde trenler kısmi çizildi.\n\n" +
                    string.Join(Environment.NewLine, gr),
                    "Eksik runtime uyarısı",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                _warnedMissingRt = true;
            }


            var viols = DetectShiftViolations();
            var violSet = new HashSet<string>(viols.Select(v => v.tid));
            _plot.ClearAll();

            // seçili tren varsa turuncunun korunması için id’yi al
            string selectedTid = null;
            if (lstTrains.SelectedIndex >= 0)
                selectedTid = _engine.Trains.Keys.OrderBy(k => k).ElementAt(lstTrains.SelectedIndex);

            var head = GetHead(); var clear = GetClear(); var slack = GetSlack();
            var (confs, overlays) = BuildOverlays(head, clear, slack);
            var plan = BuildPlacementPlan(head, clear, slack);
            _unscheduled.Clear();
            foreach (var u in plan.skipped) _unscheduled.Add(u);

            // eskisi: var name2 = _engine.Stations.ToDictionary(s => s.Name, s => (s.Name, s.Km));
            var name2 = _engine.Stations.ToDictionary(
                s => s.Name,
                s => (s.Name, s.Km),
                StringComparer.OrdinalIgnoreCase
            );


            // mor uyarı listesi
            var warn = new List<string>();

            foreach (var tid in plan.placed)
            {
                var t = _engine.Trains[tid];

                // SADECE hesaplanmış kısmı çiz
                var visibleNames = TrimToComputed(tid);
                if (visibleNames.Count == 0) continue; // bu treni pas geç (eksik runtime yüzünden hiç hesap yok)

                // name2’de olmayan istasyonları at
                var pathSt = visibleNames
                    .Where(n => name2.ContainsKey(n))
                    .Select(n => name2[n])
                    .ToList();

                bool isDashed = plan.dashFrom.ContainsKey(tid);
                var color = isDashed ? OxyColor.FromRgb(255, 0, 0) : ColorByType(t.Type);
                bool dashed = isDashed;

                if (ExceedsShiftLimit(tid, out int deltaAbs, out int limit))
                {
                    color = Violet;
                    string sign = (_engine.Trains[tid].ReqTime - _baseReq[tid]) >= 0 ? "+" : "-";
                    warn.Add($"{tid} [{t.Type}] {sign}{deltaAbs} dk (> {limit} dk) — feasible, OPT değil");
                }

                // overlay’leri de görünür path’e göre filtrele
                if (!overlays.TryGetValue(tid, out var ov))
                    ov = new List<((string u, string v) seg, int s, int e)>();
                var setVis = new HashSet<string>(visibleNames);
                ov = ov.Where(x => setVis.Contains(x.seg.u) && setVis.Contains(x.seg.v)).ToList();

                _plot.DrawTrain(tid, pathSt, _engine.Times[tid], t.Dwell, color, ov, dashed, null);
            }


            // full-line kırmızı (yalnız planlı-planlı çakışanlar)
            var setUns = new HashSet<string>(_unscheduled);
            var full = new HashSet<string>();
            foreach (var c in confs)
                if (!setUns.Contains(c.A) && !setUns.Contains(c.B))
                { full.Add(c.A); full.Add(c.B); }
            _plot.HighlightTrains(full);
            _plot.MarkUnscheduled(_unscheduled);

            // seçili tren turuncu kalsın
            if (!string.IsNullOrEmpty(selectedTid))
                _plot.SetSelectedTrain(selectedTid);

            // metin özet
            var lines = new List<string>();
            foreach (var c in confs)
                lines.Add($"{c.A} ↔ {c.B} @ {c.Seg.u}-{c.Seg.v}  [{c.Interval.s}-{c.Interval.e}]  {(c.SameDir ? "same-dir" : "opp-dir")}");
            if (_unscheduled.Any())
            {
                var pgs = _unscheduled.Where(tid => _engine.Trains[tid].Type.ToUpper() != "F").OrderBy(z => z);
                var fs = _unscheduled.Where(tid => _engine.Trains[tid].Type.ToUpper() == "F").OrderBy(z => z);
                lines.Add("");
                if (pgs.Any()) lines.Add("Yerleştirilemeyen P/YHT: " + string.Join(", ", pgs));
                if (fs.Any()) lines.Add("Yerleştirilemeyen Yük: " + string.Join(", ", fs));
            }
            txtConfMain.Text = lines.Count == 0 ? "Çakışma yok." : string.Join(Environment.NewLine, lines);

            _lastConflicts = confs;
            _lastShiftWarnings = warn;

            // mor uyarı penceresini güncelle
            UpdateWarningsWindow(viols);
        }

        // ==== BUTONLAR ====
        private void BtnReload_Click(object sender, RoutedEventArgs e) => LoadCsvDataAndRender();
        // PDF olarak kaydet (OxyPlot.SkiaSharp)
        // ======== EXPORT (PNG/PDF) ========

        // Dinamik çıktı boyutu (istasyon sayısı + ekranda görünen zaman aralığı)
        private (int w, int h) ComputeExportSize()
        {
            int stCount = _engine?.Stations?.Count ?? 10;
            int h = Math.Max(700, stCount * 40);  // istasyon başına ~40px

            var xAxis = _plot?.Model?.Axes?.FirstOrDefault(a => a.Position == OxyPlot.Axes.AxisPosition.Bottom)
                        as OxyPlot.Axes.LinearAxis;
            double spanMin = xAxis != null ? (xAxis.ActualMaximum - xAxis.ActualMinimum) : 12 * 60; // 12 saat varsayılan
            int w = (int)Math.Max(1200, (spanMin / 60.0) * 160); // saat başına ~160px

            return (w, h);
        }

        // PNG: model tabanlı export (kırpma yok)
        private void SavePlotAsPng_Exporter(string path, int width, int height, int dpi = 96)
        {
            // WPF PngExporter (senin sürümünde Width/Height var; Dpi/Background yok)
            var exporter = new OxyPlot.Wpf.PngExporter
            {
                Width = width,
                Height = height
            };

            using (var s = System.IO.File.Create(path))
            {
                exporter.Export(_plot.Model, s);
            }
        }

        // PDF: OxyPlot.Pdf (Skia yok, native bağımlılık yok)
        // PDF (OxyPlot.PdfExporter; Skia yok)
        private void SavePlotAsPdf_Managed(string path, int width, int height)
        {
            // PDF exporter’da Background property yok → Model.Background beyazsa yeter.
            var oldBg = _plot.Model.Background;
            _plot.Model.Background = OxyColors.White;   // garanti olsun

            try
            {
                var exporter = new PdfExporterManaged
                {
                    Width = width,
                    Height = height
                };
                using (var s = System.IO.File.Create(path))
                    exporter.Export(_plot.Model, s);
            }
            finally
            {
                _plot.Model.Background = oldBg;
            }
        }




        private void BtnSavePng_Click(object sender, RoutedEventArgs e)
        {
            var (w, h) = ComputeExportSize();

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Grafiği Kaydet",
                Filter = "PNG Görseli (*.png)|*.png|PDF Belgesi (*.pdf)|*.pdf",
                FileName = $"schedule_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (dlg.ShowDialog() == true)
            {
                var ext = System.IO.Path.GetExtension(dlg.FileName).ToLowerInvariant();
                try
                {
                    if (ext == ".pdf")
                        SavePlotAsPdf_Managed(dlg.FileName, w, h);
                    else
                        SavePlotAsPng_Exporter(dlg.FileName, w, h, dpi: 96);

                    MessageBox.Show("Kaydedildi:\n" + dlg.FileName, "Kaydet",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kaydetme hatası:\n" + ex.Message, "Hata",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }



        /// <summary>
        /// plotView kontrolünü olduğu gibi PNG’ye render eder.
        /// </summary>


        private void BtnConf_Click(object sender, RoutedEventArgs e)
        {
            RefreshPlot();
            MessageBox.Show(_lastConflicts.Count == 0 ? "Çakışma yok." : $"Çakışma sayısı: {_lastConflicts.Count}", "Çakışmalar");
        }

        private void BtnShiftMinus_Click(object sender, RoutedEventArgs e) => ShiftSelected(-Math.Abs(ParseInt(txtDelta.Text, 10)));
        private void BtnShiftPlus_Click(object sender, RoutedEventArgs e) => ShiftSelected(+Math.Abs(ParseInt(txtDelta.Text, 10)));

        private void ShiftSelected(int delta)
        {
            if (lstTrains.SelectedIndex < 0) { MessageBox.Show("Önce bir tren seçin."); return; }
            var tid = _engine.Trains.Keys.OrderBy(k => k).ElementAt(lstTrains.SelectedIndex);
            var t = _engine.Trains[tid];
            t.ReqTime = Math.Max(0, Math.Min(1440, t.ReqTime + delta));
            _engine.ComputeTimes();

            // listbox metinlerini tazele
            lstTrains.Items.Clear();
            foreach (var tr in _engine.Trains.Values.OrderBy(z => z.ReqTime))
                lstTrains.Items.Add($"{tr.Id} [{tr.Type},{tr.Direction}] {tr.Origin}->{(string.IsNullOrWhiteSpace(tr.Dest) ? _engine.Paths[tr.Id].Last() : tr.Dest)} @{tr.ReqTime}m");

            RefreshPlot();

            // manuel kaydırmadan sonra mor uyarıları göster
            if (_lastShiftWarnings.Any())
                MessageBox.Show("Mor çizilen trenler kaydırma sınırını aşmıştır (feasible, OPT değil):\n\n" +
                                string.Join(Environment.NewLine, _lastShiftWarnings),
                                "Kaydırma Sınırı Uyarısı", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void BtnApplyDwell_Click(object sender, RoutedEventArgs e)
        {
            if (lstTrains.SelectedIndex < 0) { MessageBox.Show("Önce bir tren seçin."); return; }
            var tid = _engine.Trains.Keys.OrderBy(k => k).ElementAt(lstTrains.SelectedIndex);
            var st = (cbStation.SelectedItem as string) ?? "";
            int d = ParseInt(txtDwell.Text, 2);
            if (string.IsNullOrWhiteSpace(st)) return;
            if (!_engine.Trains[tid].Dwell.ContainsKey(st)) _engine.Trains[tid].Dwell[st] = 0;
            _engine.Trains[tid].Dwell[st] = d;
            RefreshPlot();
        }

        // ==== OxyPlot pan/zoom ====
        private void PanLeft_Click(object s, RoutedEventArgs e) => _plot.PanX(-30);
        private void PanRight_Click(object s, RoutedEventArgs e) => _plot.PanX(+30);
        private void PanUp_Click(object s, RoutedEventArgs e) => _plot.PanY(+5);
        private void PanDown_Click(object s, RoutedEventArgs e) => _plot.PanY(-5);
        private void ZoomIn_Click(object s, RoutedEventArgs e) => _plot.Zoom(0.8);
        private void ZoomOut_Click(object s, RoutedEventArgs e) => _plot.Zoom(1.25);

        // *** ÖNEMLİ: TimeSpanAxis kullandığımız için dakika ile ayarla ***
        private void ResetView_Click(object s, RoutedEventArgs e) => _plot.ResetView12h();

        // ==== Gelişmiş diyalog ====
        private AdvancedWindow _adv;

        private void BtnAdvanced_Click(object sender, RoutedEventArgs e)
        {
            if (_adv == null || !_adv.IsLoaded)
            {
                _adv = new AdvancedWindow(this);
                _adv.Owner = this;
                // Her ihtimale karşı: biri koddan Close() derse referansı sıfırla
                _adv.Closed += (_, __) => _adv = null;
            }

            if (!_adv.IsVisible) _adv.Show();
            if (_adv.WindowState == WindowState.Minimized)
                _adv.WindowState = WindowState.Normal;

            _adv.Activate();
            _adv.Focus();

            _adv.SyncStateFromMain();
        }

        // ——— ÖNERİLER: ekranda göster + "Uygulansın mı?" sor
        private bool ShowSuggestionsAskApply(IEnumerable<string> lines, string filePath, string title)
        {
            var arr = lines?.ToList() ?? new List<string>();
            int preview = Math.Min(15, arr.Count);
            string msg = (preview > 0)
                ? string.Join(Environment.NewLine, arr.Take(preview))
                : "Öneri bulunamadı.";

            if (arr.Count > preview)
                msg += $"\n... (+{arr.Count - preview} öneri daha)";

            if (!string.IsNullOrWhiteSpace(filePath))
                msg += $"\n\nTam liste: {filePath}";

            msg += "\n\nBu önerileri otomatik uygulayalım mı?";

            var res = MessageBox.Show(msg, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            return res == MessageBoxResult.Yes;
        }

        // ——— Bir tren, bu fiziksel segmentte içeri girerken bekletilecek istasyon (yukarı akış istasyonu)
        private string UpstreamOnSegment(string tid, (string a, string b) phys)
        {
            var path = _engine.Paths[tid];
            for (int i = 0; i + 1 < path.Count; i++)
            {
                if (path[i] == phys.a && path[i + 1] == phys.b) return phys.a;
                if (path[i] == phys.b && path[i + 1] == phys.a) return phys.b;
            }
            return phys.a; // bulunamazsa güvenli
        }

        // ——— Bir çatışma için "duruş öncelikli" en küçük düzeltmeyi hesapla.
        private (string tid, string station, int addDwell, int addShift) ComputeGreedyFix(
            Conflict c, int headSame, int clearOpp, int stationSlack, int dwellCapPerStation = 20)
        {
            var phys = PhysKey(c.Seg.u, c.Seg.v);

            (int s, int e) Win(string tid)
            {
                var win = _engine.SegmentWindows(tid).FirstOrDefault(w => PhysKey(w.seg.u, w.seg.v) == phys);
                if (!(win.seg.u is null) && !(win.seg.v is null)) return (win.s, win.e);

                var path = _engine.Paths[tid];
                int ix = path.IndexOf(phys.a);
                if (ix >= 0 && ix + 1 < path.Count && path[ix + 1] == phys.b)
                {
                    int aA = _engine.Times[tid][phys.a];
                    int dwA = _engine.Trains[tid].Dwell.TryGetValue(phys.a, out var dA) ? dA : 0;
                    int dA2 = aA + dwA;
                    int aB = _engine.Times[tid][phys.b];
                    return (dA2, aB);
                }
                ix = path.IndexOf(phys.b);
                if (ix >= 0 && ix + 1 < path.Count && path[ix + 1] == phys.a)
                {
                    int aB = _engine.Times[tid][phys.b];
                    int dwB = _engine.Trains[tid].Dwell.TryGetValue(phys.b, out var dB) ? dB : 0;
                    int dB2 = aB + dwB;
                    int aA = _engine.Times[tid][phys.a];
                    return (dB2, aA);
                }
                return (c.Interval.s, c.Interval.e);
            }

            var (sA, eA) = Win(c.A);
            var (sB, eB) = Win(c.B);

            int prA = TrainPriority(_engine.Trains[c.A]);
            int prB = TrainPriority(_engine.Trains[c.B]);

            bool same = c.SameDir;
            string delay = c.B;
            int need = 0;

            if (same)
            {
                bool aLead = sA <= sB;
                int leadIn = aLead ? sA : sB;
                int leadOut = aLead ? eA : eB;
                int folIn = aLead ? sB : sA;
                int folOut = aLead ? eB : eA;

                int d1 = headSame - (folIn - leadIn);
                int d2 = headSame - (folOut - leadOut);
                need = Math.Max(0, Math.Max(d1, d2));

                if (prA > prB) delay = c.A;
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }
            else
            {
                if (eA <= sB) need = Math.Max(0, clearOpp - (sB - eA));
                else need = Math.Max(0, clearOpp - (sA - eB));

                if (prA > prB) delay = c.A;
                else if (prB > prA) delay = c.B;
                else delay = (_engine.Trains[c.A].ReqTime <= _engine.Trains[c.B].ReqTime ? c.B : c.A);
            }

            if (need <= 0) need = 1; // emniyet

            string st = UpstreamOnSegment(delay, PhysKey(c.Seg.u, c.Seg.v)); // bekletilecek istasyon
            int baseDw = _engine.Trains[delay].Dwell.TryGetValue(st, out var d0) ? d0 : 0;
            int room = Math.Max(0, dwellCapPerStation - baseDw);
            int addDw = Math.Min(room, need);
            int addSh = need - addDw;

            return (delay, st, addDw, addSh);
        }

        // ——— Greedy: tek istasyon ÇAPA ile (merdiveni önler)
        private (int steps, int left, List<string> log) AutoResolveByGreedyDurusFirst(
            int headSame, int clearOpp, int stationSlack, int maxIters = 500, int dwellCapPerStation = MAX_DWELL_PER_STATION)
        {
            var log = new List<string>();
            var anchorOf = new Dictionary<string, string>(); // her tren için tek bekleme istasyonu
            int steps = 0;

            for (int it = 0; it < maxIters; it++)
            {
                var confs = _engine.DetectConflicts(headSame, clearOpp, stationSlack);
                if (confs.Count == 0) return (steps, 0, log);

                // En erken çatışma ve toplam ihtiyaç
                var c = confs.OrderBy(z => z.Interval.s).First();
                var fix = ComputeGreedyFix(c, headSame, clearOpp, stationSlack, dwellCapPerStation);
                int need = fix.addDwell + fix.addShift;

                // Bu tren için çapa istasyonu seç/hatırla
                if (!anchorOf.TryGetValue(fix.tid, out var anchor) || string.IsNullOrEmpty(anchor))
                {
                    anchor = string.IsNullOrEmpty(fix.station)
                           ? _engine.Paths[fix.tid].First()  // yedek
                           : fix.station;
                    anchorOf[fix.tid] = anchor;
                }

                var t = _engine.Trains[fix.tid];

                // Önce çapada biriktir
                int baseDw = t.Dwell.TryGetValue(anchor, out var d0) ? d0 : 0;
                int room = Math.Max(0, dwellCapPerStation - baseDw);
                int addAtAnchor = Math.Min(room, need);
                if (addAtAnchor > 0)
                {
                    t.Dwell[anchor] = baseDw + addAtAnchor;
                    log.Add($"{fix.tid} @ {anchor} +{addAtAnchor} dk duruş");
                }

                // Kalanı başlangıç saatini kaydır
                int remain = need - addAtAnchor;
                if (remain > 0)
                {
                    t.ReqTime = Math.Max(0, Math.Min(1440, t.ReqTime + remain));
                    log.Add($"{fix.tid} +{remain} dk kaydırma");
                }

                _engine.ComputeTimes();
                steps++;
            }

            return (steps, _engine.DetectConflicts(headSame, clearOpp, stationSlack).Count, log);
        }

        // ==== Optimize (C# / OR-Tools) ====
        private void BtnOptimize_Click(object sender, RoutedEventArgs e)
        {
            if (_engine.LastMissingRuntimes.Count > 0)
            {
                MessageBox.Show(
                    "Eksik runtime varken optimizasyon çalıştırılamaz.\n" +
                    "Lütfen 'runtimes' sayfasında bütün (u,v,type) sürelerini girin.",
                    "Optimizasyon engellendi", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            var swTotal = Stopwatch.StartNew();

            // OR-Tools hazır mı?
            try
            {
                NativeOrtools.EnsureLoaded();
                var m = new Google.OrTools.Sat.CpModel();
                m.NewBoolVar("x");
                var s = new Google.OrTools.Sat.CpSolver();
                s.Solve(m);
            }
            catch (Exception ex)
            {
                MessageBox.Show("OR-Tools yüklenemedi:\n\n" + ex, "OR-Tools", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // her tren için TEK “çapa” istasyon: ilk ara istasyon
                var allowed = new Dictionary<string, List<string>>();
                foreach (var tid in _engine.Trains.Keys)
                {
                    var seq = _engine.Paths[tid];
                    string anchor = (seq.Count >= 2) ? seq[1] : seq[0]; // tek çapa
                    allowed[tid] = new List<string> { anchor };
                }

                var swSolve = Stopwatch.StartNew();
                var (shifts, status, placed, dropped) = OptimizerCs.OptimizeSchedule(
                    _engine,
                    headSame: GetHead(),
                    clearOpp: GetClear(),
                    stationSlack: GetSlack(),
                    nodeSep: 1,
                    dwellMax: 15,
                    timeLimit: 10,
                    // dwell kaydırmadan daha ucuz (tercih edilir)
                    wTotalDwell: 5,   // dwell 1 dk = 5 puan
                    wDevstart: 50,    // shift 1 dk = 50 puan  (≈10:1)
                    wSpan: 1,
                    allowedAutoDwell: allowed,
                    allowDrop: true,
                    wPlace: 1_000_000,
                    wCompact: 0,
                    dwellCapPerStation: 20
                );
                swSolve.Stop();

                // çözüm uygulandı → çiz
                RefreshPlot();

                // CP-SAT 'Unknown' gelse bile çözüm değişkenlerde; yalnız placed==0 ise öneriye geç
                bool ok = status == Google.OrTools.Sat.CpSolverStatus.Optimal
                       || status == Google.OrTools.Sat.CpSolverStatus.Feasible
                       || status == Google.OrTools.Sat.CpSolverStatus.Unknown;

                if (!ok || placed.Count == 0)
                {
                    var head = GetHead(); var clear = GetClear(); var slack = GetSlack();
                    var sug = MakeSuggestions(head, clear, slack);
                    var file = WriteSuggestionsToDesktop(sug);
                    bool apply = ShowSuggestionsAskApply(sug, file, "Optimizer Önerileri");
                    if (apply)
                    {
                        var (steps, left, _) = AutoResolveByGreedyDurusFirst(head, clear, slack, maxIters: 1000, dwellCapPerStation: MAX_DWELL_PER_STATION);
                        RefreshPlot();
                        MessageBox.Show($"Öneriler uygulandı. Adım: {steps}, Kalan: {left}", "Optimizer");
                    }
                    return;
                }

                // hâlâ çatışma kaldıysa: önce otomatik yerel onarım dene (sorma)
                var h = GetHead(); var c = GetClear(); var sl = GetSlack();
                var afterCp = _engine.DetectConflicts(h, c, sl).Count;
                if (afterCp > 0)
                {
                    var (steps, left, _) = AutoResolveByGreedyDurusFirst(h, c, sl, maxIters: 1000, dwellCapPerStation: MAX_DWELL_PER_STATION);
                    RefreshPlot();
                }

                // mor uyarıları göster
                if (_lastShiftWarnings.Any())
                {
                    MessageBox.Show("Mor çizilen trenler kaydırma sınırını aşmıştır (feasible, OPT değil):\n\n" +
                                    string.Join(Environment.NewLine, _lastShiftWarnings),
                                    "Kaydırma Sınırı Uyarısı", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show($"C# Optimizer tamamlandı.\nÇözüm süresi: {swSolve.Elapsed.TotalSeconds:0.00} sn", "Optimizer");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Optimizer Hata");
            }
            finally
            {
                swTotal.Stop();
            }
        }

        // gelişmiş pencereden okunan tamponlar (yoksa default)
        internal int GetHead() => _adv?.Head ?? 3;
        internal int GetClear() => _adv?.Clear ?? 3;
        internal int GetSlack() => _adv?.Slack ?? 1;

        // export helpers
        internal void ExportTimetableCsv(string path, IList<string> onlyStations, bool outMin, bool outHHMM)
        {
            var order = _engine.Stations.Select(s => s.Name).ToList();
            var types = _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Type);
            var dwells = _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Dwell);
            ExcelXlsx.WriteTimetableCsv(path, order, _engine.Times, types, dwells, onlyStations?.ToList(), outMin, outHHMM);
        }

        internal void SaveExcelTemplate(string path) => ExcelXlsx.WriteTemplate(path);

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Timetable CSV kaydet",
                Filter = "CSV dosyası (*.csv)|*.csv",
                FileName = "timetable.csv"
            };
            if (dlg.ShowDialog() != true) return;

            ExportTimetableCsv(
                dlg.FileName,
                onlyStations: null,   // veya bir filtre listesi ver
                outMin: false,
                outHHMM: true
            );
            MessageBox.Show("CSV dışa aktarıldı:\n" + dlg.FileName, "CSV", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        internal void SaveExcelDataset(string path)
        {
            var types = _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Type);
            var dwells = _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Dwell);

            // Runtime listesi — Engine.Run sözlüğünden yeniden kuruyoruz
            IList<Runtime> runtimesForXlsx = _engine?.Run?
                .Select(kv => new Runtime(
                    kv.Key.u,          // u
                    kv.Key.v,          // v
                    kv.Key.tt,         // type  (P/F; G -> F map’lenmiş halde)
                    kv.Value           // run_min
                ))
                .ToList();

            // … ve burada null yerine runtimesForXlsx gönder
            ExcelXlsx.WriteDatasetXlsx(
                path,
                _engine.Stations.ToList(),
                _engine.Segments.ToList(),
                _engine.Trains.Values.ToList(),
                runtimesForXlsx,   // <— burası
                _engine.Times,
                _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Type),
                _engine.Trains.ToDictionary(kv => kv.Key, kv => kv.Value.Dwell),
                headSame: GetHead(),
                clearOpp: GetClear(),
                stationSlack: GetSlack()
            );

            // Eski:
            // ExcelXlsx.WriteDataset(path, ...);
            MessageBox.Show("Excel dataset kaydedildi:\n" + path, "Excel");
        }

        // mevcut saat - referans saat farklarına göre morlukları bul
        private List<(string tid, string typedesc, int delta, int limit)> DetectShiftViolations()
        {
            var list = new List<(string, string, int, int)>();
            foreach (var kv in _engine.Trains)
            {
                var t = kv.Value;
                if (!_baseReq.TryGetValue(kv.Key, out var baseStart)) continue;
                int delta = t.ReqTime - baseStart;
                int lim = ShiftLimit(t);
                if (Math.Abs(delta) > lim)
                {
                    string typeDesc = $"{t.Type}";
                    list.Add((kv.Key, typeDesc, delta, lim));
                }
            }
            return list;
        }

        // mor uyarı penceresini aç/kapat/güncelle
        private void UpdateWarningsWindow(List<(string tid, string typedesc, int delta, int limit)> viols)
        {
            if (viols != null && viols.Count > 0)
            {
                if (_warnWin == null || !_warnWin.IsLoaded)
                {
                    _warnWin = new WarningsWindow();
                    _warnWin.Owner = this;

                    // buton event’lerine abone ol
                    _warnWin.AcceptAllRequested += () =>
                    {
                        // o anki morların hepsini referans kabul et
                        var current = DetectShiftViolations();
                        foreach (var v in current)
                            if (_engine.Trains.TryGetValue(v.tid, out var t))
                                _baseReq[v.tid] = t.ReqTime;

                        RefreshPlot(); // morlar silinir, pencere de kapanır
                    };

                    _warnWin.AcceptSelectedRequested += (tid) =>
                    {
                        if (string.IsNullOrWhiteSpace(tid)) return;
                        if (_engine.Trains.TryGetValue(tid, out var t))
                        {
                            _baseReq[tid] = t.ReqTime;
                            RefreshPlot();
                        }
                    };

                    // sağ-üstte dursun (mevcut konumlandırmanla aynı mantık)
                    _warnWin.Left = this.Left + this.Width - _warnWin.Width - 24;
                    _warnWin.Top = this.Top + 64;
                    _warnWin.Show();
                }

                var lines = viols
                    .OrderBy(v => v.tid)
                    .Select(v =>
                    {
                        string sign = v.delta >= 0 ? "+" : "";
                        return $"{v.tid} [{v.typedesc}] {sign}{v.delta} dk (> {v.limit} dk) — feasible, OPT değil";
                    });

                _warnWin.SetItems(lines);
                if (!_warnWin.IsVisible) _warnWin.Show();
                _warnWin.Activate();
            }
            else
            {
                // mor kalmadıysa pencereyi kapat
                if (_warnWin != null)
                {
                    _warnWin.Close();
                    _warnWin = null;
                }
            }
        }


        internal void LoadExcelDataset(string path)
        {
            // .xlsx ise doğrudan Excel’den, değilse CSV setinden (stem_* dosyaları) oku
            List<Station> stations;
            List<(string u, string v)> segments;
            List<Train> trains;
            List<Runtime> runtimes;

            var ext = System.IO.Path.GetExtension(path);
            if (ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                (stations, segments, trains, runtimes) = ExcelXlsx.ReadDatasetXlsx(path);
            }
            else
            {
                (stations, segments, trains, runtimes) = ExcelXlsx.ReadDataset(path);
            }

            _engine = new Engine(stations, segments, trains, runtimes);
            _engine.ComputeTimes();

            // referans başlangıç saatlerini mor limit kontrolü için kaydet
            _baseReq.Clear();
            foreach (var kv in _engine.Trains)
                _baseReq[kv.Key] = kv.Value.ReqTime;

            // CSV’den gelmişse duruşları kapla (Excel için de sorun olmaz)
            ClampAllDwellsToCap();

            lstTrains.Items.Clear();
            foreach (var t in _engine.Trains.Values.OrderBy(t => t.ReqTime))
                lstTrains.Items.Add($"{t.Id} [{t.Type},{t.Direction}] {t.Origin}->{(string.IsNullOrWhiteSpace(t.Dest) ? _engine.Paths[t.Id].Last() : t.Dest)} @{t.ReqTime}m");

            cbStation.Items.Clear();
            foreach (var s in _engine.Stations) cbStation.Items.Add(s.Name);
            if (cbStation.Items.Count > 0) cbStation.SelectedIndex = 0;

            _unscheduled.Clear();
            _plot.ResetStations(_engine.Stations.Select(s => (s.Name, s.Km)));
            RefreshPlot();
            CheckStationConsistency();

            _warnedMissingRt = false;

        }
        private void CheckStationConsistency()
        {
            var inStations = _engine.Stations.Select(s => s.Name.Trim())
                              .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var inSegments = _engine.Segments
                              .SelectMany(e => new[] { e.u.Trim(), e.v.Trim() })
                              .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var inTrains = _engine.Trains.Values
                              .SelectMany(t => new[] { t.Origin?.Trim(), t.Dest?.Trim() })
                              .Where(s => !string.IsNullOrWhiteSpace(s))
                              .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var missing = inSegments.Union(inTrains)
                            .Where(n => !inStations.Contains(n))
                            .OrderBy(n => n).ToList();

            if (missing.Count > 0)
                MessageBox.Show("stations sayfasında eksik olan istasyonlar:\n\n" +
                                string.Join("\n", missing),
                                "Eksik İstasyon(lar)");
        }


        private static int ParseInt(string s, int def) => int.TryParse(s, out var v) ? v : def;

        // AdvancedWindow'dan tetiklemek için:
        public void InvokeCheckConflicts() => BtnConf_Click(this, new RoutedEventArgs());
        public void InvokeOptimize() => BtnOptimize_Click(this, new RoutedEventArgs());
        public void InvokeRefreshPlot() => RefreshPlot();
        public IEnumerable<string> StationNames() =>
            _engine?.Stations.Select(s => s.Name) ?? Enumerable.Empty<string>();
        public string ConflictSummaryText() => txtConfMain?.Text ?? "";
        private static string HHMM(int m)
        {
            m = ((m % 1440) + 1440) % 1440;
            return $"{(m / 60) % 24:00}:{m % 60:00}";
        }
    }
}
