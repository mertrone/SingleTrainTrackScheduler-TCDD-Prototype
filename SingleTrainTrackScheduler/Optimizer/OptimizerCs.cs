using Google.OrTools.Sat;
using SingleTrainTrackScheduler.Core;
using SingleTrainTrackScheduler.util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SingleTrainTrackScheduler.Optimizer
{
    public static class OptimizerCs
    {
        public static (Dictionary<string, int> shifts,
                       CpSolverStatus status,
                       List<string> placed,
                       List<string> dropped)
        OptimizeSchedule(
            Engine engine,
            int headSame,
            int clearOpp,
            int stationSlack = 0,
            int nodeSep = 1,
            int dwellMax = 15,
            int timeLimit = 10,
            int wTotalDwell = 5000,         // dwell çok pahalı
            int wDevstart = 10,             // kaydırma daha ucuz
            int wSpan = 1,
            Dictionary<string, List<string>> allowedAutoDwell = null,
            bool allowDrop = true,
            int wPlace = 1_000_000,
            int wCompact = 0,               // gereksiz dwell’i tetiklemesin
            int dwellCapPerStation = 20
        )
        {
            engine.ComputeTimes();

            var model = new CpModel();
            var tids = engine.Trains.Keys.OrderBy(k => k).ToList();
            var stations = engine.Stations.Select(s => s.Name).ToList();
            var stIndex = engine.StationIndex;
            const int BIG = 10_000_000;

            // ---------- taban zamanlar ----------
            var arrBase = new Dictionary<string, Dictionary<string, int>>();
            var depBase = new Dictionary<string, Dictionary<string, int>>();
            var fixedDw = new Dictionary<string, Dictionary<string, int>>();

            foreach (var tid in tids)
            {
                var t = engine.Trains[tid];
                fixedDw[tid] = new Dictionary<string, int>(t.Dwell ?? new Dictionary<string, int>());
                arrBase[tid] = new Dictionary<string, int>();
                depBase[tid] = new Dictionary<string, int>();
                foreach (var s in engine.Paths[tid])
                {
                    if (!engine.Times[tid].ContainsKey(s)) continue;
                    int a = engine.Times[tid][s];
                    int d = a + (fixedDw[tid].TryGetValue(s, out var dw) ? dw : 0);
                    arrBase[tid][s] = a;
                    depBase[tid][s] = d;
                }
            }

            // ---------- karar değişkenleri ----------
            var x = new Dictionary<string, IntVar>();      // başlangıç kaydırması
            var absX = new Dictionary<string, IntVar>();   // |x|
            foreach (var tid in tids)
            {
                var tt = (engine.Trains[tid].Type ?? "").ToUpperInvariant();
                int maxDev = (tt == "P" || tt == "G") ? 15 : 45; // Yük/YHT sınırları
                var xv = model.NewIntVar(-maxDev, maxDev, $"x_{tid}");
                var ax = model.NewIntVar(0, maxDev, $"abs_x_{tid}");
                model.AddAbsEquality(ax, xv);
                x[tid] = xv; absX[tid] = ax;
            }

            // opsiyonel aktiflik
            var active = new Dictionary<string, BoolVar>();
            if (allowDrop) foreach (var tid in tids) active[tid] = model.NewBoolVar($"active_{tid}");

            // izinli ek dwell (tek/az istasyon)
            allowedAutoDwell ??= new Dictionary<string, List<string>>();
            var dExtra = new Dictionary<(string tid, string s), IntVar>();
            foreach (var tid in tids)
            {
                var allow = (allowedAutoDwell.TryGetValue(tid, out var lst) && lst != null)
                            ? lst.Distinct().ToList() : new List<string>();
                foreach (var s in allow)
                {
                    int baseDw = fixedDw.TryGetValue(tid, out var map) && map.TryGetValue(s, out var d0) ? d0 : 0;
                    int cap = Math.Max(0, dwellCapPerStation - baseDw);
                    int ub = Math.Min(dwellMax, cap);
                    var v = model.NewIntVar(0, ub, $"dwell_extra_{tid}_{s}");
                    model.Add(v <= cap);
                    dExtra[(tid, s)] = v;
                }
            }

            // yardımcı ifadeler
            LinearExpr ARR(string tid, string s) => x[tid] + arrBase[tid][s];
            LinearExpr DEP(string tid, string s)
            {
                LinearExpr expr = x[tid] + depBase[tid][s];
                if (dExtra.TryGetValue((tid, s), out var dv)) expr += dv;
                return expr;
            }

            (string a, string b) Phys(string u, string v) => stIndex[u] < stIndex[v] ? (u, v) : (v, u);

            // ---------- KENDİNİ SAĞLAMA (her tren, her kenar) ----------
            foreach (var tid in tids)
            {
                var path = engine.Paths[tid];
                for (int i = 0; i + 1 < path.Count; i++)
                {
                    var u = path[i]; var v = path[i + 1];
                    if (!arrBase[tid].ContainsKey(u) || !arrBase[tid].ContainsKey(v)) continue;

                    // ARR(v) - slack >= DEP(u) + slack
                    if (allowDrop)
                        model.Add(ARR(tid, v) - stationSlack >= DEP(tid, u) + stationSlack)
                             .OnlyEnforceIf(active[tid]);
                    else
                        model.Add(ARR(tid, v) - stationSlack >= DEP(tid, u) + stationSlack);
                }
            }

            // ---------- HEADWAY: aynı fiziksel segmentteki tren çiftleri ----------
            var segToItems = new Dictionary<(string a, string b), List<(string tid, string u, string v)>>();
            foreach (var tid in tids)
            {
                var path = engine.Paths[tid];
                for (int i = 0; i + 1 < path.Count; i++)
                {
                    var u = path[i]; var v = path[i + 1];
                    if (!arrBase[tid].ContainsKey(u) || !arrBase[tid].ContainsKey(v)) continue;
                    var pseg = Phys(u, v);
                    segToItems.TryAdd(pseg, new List<(string, string, string)>());
                    segToItems[pseg].Add((tid, u, v));
                }
            }

            foreach (var kv in segToItems)
            {
                var items = kv.Value; int n = items.Count;
                for (int i = 0; i < n; i++)
                {
                    var (ti, ui, vi) = items[i];
                    for (int j = i + 1; j < n; j++)
                    {
                        var (tj, uj, vj) = items[j];

                        var s_i = DEP(ti, ui) + stationSlack;
                        var e_i = ARR(ti, vi) - stationSlack;
                        var s_j = DEP(tj, uj) + stationSlack;
                        var e_j = ARR(tj, vj) - stationSlack;

                        bool sameDir = engine.Trains[ti].Direction == engine.Trains[tj].Direction;
                        int gap = 2 * (sameDir ? headSame : clearOpp);

                        var y = model.NewBoolVar($"seg_{kv.Key.a}_{kv.Key.b}_{ti}_before_{tj}");

                        if (allowDrop)
                        {
                            model.Add(s_j >= e_i + gap).OnlyEnforceIf(new ILiteral[] { y, active[ti], active[tj] });
                            model.Add(s_i >= e_j + gap).OnlyEnforceIf(new ILiteral[] { y.Not(), active[ti], active[tj] });
                        }
                        else
                        {
                            model.Add(s_j >= e_i + gap).OnlyEnforceIf(y);
                            model.Add(s_i >= e_j + gap).OnlyEnforceIf(y.Not());
                        }
                    }
                }
            }

            // ---------- NODE SEPARATION (opsiyonel boşluk küçültme değişkenleri) ----------
            var compactGaps = new List<IntVar>();
            void SepPair(string ta, string tb, LinearExpr A, LinearExpr B, int sep, string name)
            {
                var b = model.NewBoolVar(name);
                if (allowDrop)
                {
                    model.Add(B >= A + sep).OnlyEnforceIf(new ILiteral[] { b, active[ta], active[tb] });
                    model.Add(A >= B + sep).OnlyEnforceIf(new ILiteral[] { b.Not(), active[ta], active[tb] });

                    var gap = model.NewIntVar(0, BIG, $"gap_{name}");
                    model.Add(gap >= B - A - sep).OnlyEnforceIf(new ILiteral[] { b, active[ta], active[tb] });
                    model.Add(gap >= A - B - sep).OnlyEnforceIf(new ILiteral[] { b.Not(), active[ta], active[tb] });
                    compactGaps.Add(gap);
                }
                else
                {
                    model.Add(B >= A + sep).OnlyEnforceIf(b);
                    model.Add(A >= B + sep).OnlyEnforceIf(b.Not());

                    var gap = model.NewIntVar(0, BIG, $"gap_{name}");
                    model.Add(gap >= B - A - sep).OnlyEnforceIf(b);
                    model.Add(gap >= A - B - sep).OnlyEnforceIf(b.Not());
                    compactGaps.Add(gap);
                }
            }

            foreach (var s in stations)
            {
                var visit = tids.Where(tid => arrBase[tid].ContainsKey(s)).ToList();
                for (int a = 0; a < visit.Count; a++)
                {
                    var ta = visit[a];
                    for (int b = a + 1; b < visit.Count; b++)
                    {
                        var tb = visit[b];
                        SepPair(ta, tb, ARR(ta, s), ARR(tb, s), nodeSep, $"sep_arr_{s}_{ta}_{tb}");
                        SepPair(ta, tb, DEP(ta, s), DEP(tb, s), nodeSep, $"sep_dep_{s}_{ta}_{tb}");
                        SepPair(ta, tb, ARR(ta, s), DEP(tb, s), nodeSep, $"sep_mix1_{s}_{ta}_{tb}");
                        SepPair(ta, tb, DEP(ta, s), ARR(tb, s), nodeSep, $"sep_mix2_{s}_{ta}_{tb}");
                    }
                }
            }

            // ---------- MAKESPAN ----------
            var T_start = model.NewIntVar(-BIG, BIG, "T_start");
            var T_end = model.NewIntVar(-BIG, BIG, "T_end");
            foreach (var tid in tids)
            {
                var path = engine.Paths[tid];
                var first = path.First();
                var last = path.Last();
                if (arrBase[tid].ContainsKey(first))
                {
                    if (allowDrop) model.Add(T_start <= DEP(tid, first)).OnlyEnforceIf(active[tid]);
                    else model.Add(T_start <= DEP(tid, first));
                }
                if (arrBase[tid].ContainsKey(last))
                {
                    if (allowDrop) model.Add(T_end >= ARR(tid, last)).OnlyEnforceIf(active[tid]);
                    else model.Add(T_end >= ARR(tid, last));
                }

                // --- VARIŞ SAPMA KISITI (tip bazlı) ---
                int limit = 180; // default: yük
                var tt = (engine.Trains[tid].Type ?? "").ToUpperInvariant();
                if (tt is "YHT" or "Y" or "HSR" or "H") limit = 45;
                else if (tt == "P") limit = 60;

                var arrLast = ARR(tid, last);
                int baseArrLast = arrBase[tid][last];
                // |ARR_last - base| ≤ limit
                var diffUp = model.NewIntVar(-BIG, BIG, $"diff_{tid}");
                model.Add(diffUp == arrLast - baseArrLast);
                var absDiff = model.NewIntVar(0, BIG, $"abs_diff_{tid}");
                model.AddAbsEquality(absDiff, diffUp);
                if (allowDrop) model.Add(absDiff <= limit).OnlyEnforceIf(active[tid]);
                else model.Add(absDiff <= limit);
            }
            var span = model.NewIntVar(0, BIG, "span");
            model.Add(span == T_end - T_start);

            // ---------- AMAÇ ----------
            var obj = new List<LinearExpr>();
            foreach (var v in dExtra.Values) obj.Add(LinearExpr.Term(v, wTotalDwell));
            foreach (var v in absX.Values) obj.Add(LinearExpr.Term(v, wDevstart));
            obj.Add(LinearExpr.Term(span, wSpan));

            if (allowDrop) foreach (var a in active.Values) obj.Add(LinearExpr.Term(a, -wPlace));
            foreach (var g in compactGaps) obj.Add(LinearExpr.Term(g, wCompact));

            model.Minimize(LinearExpr.Sum(obj));

            // ---------- ÇÖZ ----------
            var solver = new CpSolver { StringParameters = $"num_search_workers:8,max_time_in_seconds:{timeLimit}" };
            var status = solver.Solve(model);
            if (status != CpSolverStatus.Optimal && status != CpSolverStatus.Feasible)
                return (new Dictionary<string, int>(), status, new List<string>(), tids.ToList());

            // ---------- UYGULA ----------
            var shifts = new Dictionary<string, int>();
            var placed = new List<string>();
            var dropped = new List<string>();

            foreach (var tid in tids)
            {
                if (allowDrop && solver.Value(active[tid]) == 0) { dropped.Add(tid); continue; }
                placed.Add(tid);
                int delta = (int)solver.Value(x[tid]);
                shifts[tid] = delta;
                var t = engine.Trains[tid];
                t.ReqTime = Math.Max(0, Math.Min(1440, t.ReqTime + delta));
            }

            foreach (var kv in dExtra)
            {
                var (tid, s) = kv.Key;
                if (allowDrop && placed.Contains(tid) == false) continue;
                int add = (int)solver.Value(kv.Value);
                if (add > 0)
                {
                    var t = engine.Trains[tid];
                    int baseDw = t.Dwell.TryGetValue(s, out var d) ? d : 0;
                    t.Dwell[s] = baseDw + add;
                }
            }

            engine.ComputeTimes();
            return (shifts, status, placed, dropped);
        }
    }
}
