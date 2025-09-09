using SingleTrainTrackScheduler.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SingleTrainTrackScheduler.Core
{

    public sealed class Engine
    {
        public IReadOnlyList<Station> Stations { get; private set; }
        public Dictionary<string, int> StationIndex { get; private set; } = new();
        public List<(string u, string v)> Segments { get; private set; } = new();
        public Dictionary<string, Train> Trains { get; private set; } = new();

        // run[(ttype,u,v)] = minutes ; ttype: 'P' veya 'F' (G => F gibi davranır)
        public Dictionary<(string tt, string u, string v), int> Run { get; private set; } = new();

        // path per train
        public Dictionary<string, List<string>> Paths { get; private set; } = new();

        // times[tid][station] = arr_min
        public Dictionary<string, Dictionary<string, int>> Times { get; private set; } = new();

        public Engine(IEnumerable<Station> stations,
                      IEnumerable<(string u, string v)> segments,
                      IEnumerable<Train> trains,
                      IEnumerable<Runtime> runtimes)
        {
            Reset(stations, segments, trains, runtimes);
        }

        public void Reset(IEnumerable<Station> stations,
                          IEnumerable<(string u, string v)> segments,
                          IEnumerable<Train> trains,
                          IEnumerable<Runtime> runtimes)
        {
            Stations = stations.OrderBy(s => s.Km).ToList();
            StationIndex = Stations.Select((s, i) => (s, i)).ToDictionary(t => t.s.Name, t => t.i);
            Segments = segments.ToList();
            Trains = trains.ToDictionary(t => t.Id, t => t);

            Run.Clear();
            foreach (var rt in runtimes)
            {
                var tt = (rt.TType == "F" || rt.TType == "G") ? "F" : "P";
                Run[(tt, rt.U, rt.V)] = rt.RunMin;
            }

            Paths.Clear();
            foreach (var t in Trains.Values)
                Paths[t.Id] = BuildPathForTrain(t);

            ComputeTimes();
        }

        private List<string> BuildPathForTrain(Train t)
        {
            var names = Stations.Select(s => s.Name).ToList();

            if (!StationIndex.ContainsKey(t.Origin))
                return t.Direction == "up" ? names : names.AsEnumerable().Reverse().ToList();

            if (!string.IsNullOrWhiteSpace(t.Dest) && StationIndex.ContainsKey(t.Dest) && t.Dest != t.Origin)
            {
                var io = StationIndex[t.Origin];
                var id = StationIndex[t.Dest];
                if (io < id) { t.Direction = "up"; return names.GetRange(io, id - io + 1); }
                else { t.Direction = "down"; return names.GetRange(id, io - id + 1).AsEnumerable().Reverse().ToList(); }
            }
            return t.Direction == "up" ? names : names.AsEnumerable().Reverse().ToList();
        }

        public Dictionary<string, Dictionary<string, int>> ComputeTimes()
        {
            Times = new();
            LastMissingRuntimes.Clear();

            foreach (var (tid, t) in Trains)
            {
                var path = Paths[tid];
                if (path == null || path.Count == 0) continue;

                var dict = new Dictionary<string, int> { [path[0]] = t.ReqTime };
                var tt = (t.Type == "F" || t.Type == "G") ? "F" : "P";

                // bu tren için eksik runtime görülürse buradan ileriye hesaplamayı durduracağız
                bool brokeOnMissing = false;

                for (int i = 0; i < path.Count - 1; i++)
                {
                    string u = path[i], v = path[i + 1];
                    if (!Run.TryGetValue((tt, u, v), out var rt) && !Run.TryGetValue((tt, v, u), out rt))
                    {
                        // EXCEPTION YOK — kayda al ve bu treni kısmi çiz
                        LastMissingRuntimes.Add(new MissingRuntime { Tid = tid, Type = tt, U = u, V = v });
                        brokeOnMissing = true;
                        break;
                    }

                    var dwellU = t.Dwell.TryGetValue(u, out var d) ? d : 0;
                    dict[v] = dict[u] + dwellU + rt;
                }

                Times[tid] = dict;
                // brokeOnMissing true ise bu tren yalnızca hesaplanan kısma kadar çizilecek
            }
            return Times;
        }

        public sealed class MissingRuntime
        {
            public string Tid { get; set; }
            public string Type { get; set; }   // P / F
            public string U { get; set; }
            public string V { get; set; }
        }

        public List<MissingRuntime> LastMissingRuntimes { get; private set; } = new();

        public List<((string u, string v) seg, int s, int e)> SegmentWindows(string tid)
        {
            var t = Trains[tid];
            var path = Paths[tid];
            var res = new List<((string, string), int, int)>();

            for (int i = 0; i < path.Count - 1; i++)
            {
                var u = path[i]; var v = path[i + 1];

                // zaman hesaplanmadıysa (eksik runtime) devam etmeyelim
                if (!Times[tid].ContainsKey(u) || !Times[tid].ContainsKey(v))
                    break;

                var tu = Times[tid][u];
                var tv = Times[tid][v];
                var du = t.Dwell.TryGetValue(u, out var d) ? d : 0;
                res.Add(((u, v), tu + du, tv));
            }
            return res;
        }


        public List<Conflict> DetectConflicts(int headSame = 5, int clearOpp = 3, int stationSlack = 0, double eps = 1e-3)
        {
            string Phys((string u, string v) seg)
                => StationIndex[seg.u] < StationIndex[seg.v] ? $"{seg.u}|{seg.v}" : $"{seg.v}|{seg.u}";

            var tids = Trains.Keys.ToList();
            var allWin = tids.ToDictionary(tid => tid, tid => SegmentWindows(tid));
            var confs = new List<Conflict>();

            for (int i = 0; i < tids.Count; i++)
            {
                for (int j = i + 1; j < tids.Count; j++)
                {
                    var ta = tids[i];
                    var tb = tids[j];
                    bool same = Trains[ta].Direction == Trains[tb].Direction;

                    foreach (var (segA, a1, a2) in allWin[ta])
                        foreach (var (segB, b1, b2) in allWin[tb])
                        {
                            if (Phys(segA) != Phys(segB)) continue;

                            // istasyon buffer'ı (yakın uçları kırp)
                            int ca1 = a1 + stationSlack, ca2 = a2 - stationSlack;
                            int cb1 = b1 + stationSlack, cb2 = b2 - stationSlack;
                            if (ca1 >= ca2 || cb1 >= cb2) continue;

                            if (!same)
                            {
                                // KARŞI YÖN: clearOpp kadar ayrık mı?
                                // Ayrık olma koşulu: A, B başlamadan clearOpp önce bitirir VEYA tersi
                                bool separated = (ca2 + clearOpp <= cb1 - eps) || (cb2 + clearOpp <= ca1 - eps);
                                if (!separated)
                                {
                                    var p = Phys(segA).Split('|');
                                    confs.Add(new Conflict
                                    {
                                        A = ta,
                                        B = tb,
                                        SameDir = false,
                                        Seg = (p[0], p[1]),
                                        // gösterim için gerçek örtüşen kısım
                                        Interval = (Math.Max(ca1, cb1), Math.Min(ca2, cb2))
                                    });
                                }
                            }
                            else
                            {
                                // AYNI YÖN: headway kontrolü (giriş VE çıkış)
                                // Lider/follower'ı giriş zamanına göre sırala
                                int leadIn = ca1, leadOut = ca2, folIn = cb1, folOut = cb2;
                                string lead = ta, fol = tb;

                                if (cb1 < ca1)
                                {
                                    leadIn = cb1; leadOut = cb2; folIn = ca1; folOut = ca2;
                                    lead = tb; fol = ta;
                                }

                                bool enterOk = (folIn - leadIn) >= headSame - eps;
                                bool exitOk = (folOut - leadOut) >= headSame - eps;

                                // İki eşik de sağlanmıyorsa çatışma
                                if (!(enterOk && exitOk))
                                {
                                    var p = Phys(segA).Split('|');
                                    confs.Add(new Conflict
                                    {
                                        A = lead,
                                        B = fol,
                                        SameDir = true,
                                        Seg = (p[0], p[1]),
                                        // görsel amaçlı: kesişim aralığı
                                        Interval = (Math.Max(ca1, cb1), Math.Min(ca2, cb2))
                                    });
                                }
                            }
                        }
                }
            }
            return confs;
        }

    }

    public sealed class Conflict
    {
        public string A { get; set; }
        public string B { get; set; }
        public (string u, string v) Seg { get; set; }
        public (int s, int e) Interval { get; set; }
        public bool SameDir { get; set; }
    }
}
