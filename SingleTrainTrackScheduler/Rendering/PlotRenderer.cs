using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
// kısa ad
using WpfPlotView = OxyPlot.Wpf.PlotView;

namespace SingleTrainTrackScheduler.Rendering
{
    /// <summary>
    /// Zaman ekseni dakika, Y ekseni km. İstasyon adları Y etiketlerinde.
    /// Kendi hover/tooltip’i var (OxyPlot sürümünden bağımsız).
    /// </summary>
    public sealed class PlotRenderer
    {
        private const bool ALWAYS_SHOW_SMALL_TIME_LABELS = true; // <-- ETİKETLER HER ZAMAN GÖRÜNSÜN
        private const double LABELS_DENSE_SPAN = 240.0; // dk — bu genişlikten dar ise küçük saat etiketleri görünsün

        // Hover hesaplamasını 30 ms'den sık yapma
        private readonly Stopwatch _hoverWatch = Stopwatch.StartNew();
        private long _lastHoverMs = 0;

        // Küçük saat etiketlerini sadece yeterince zoom'lu iken çiz
        private const double LABELS_DETAIL_THRESHOLD_MIN = 360; // 6 saat pencere

        // Y ekseni etiketi hesapları için
        private double _yTickStep = 5.0;

        // --- sabitler
        private const double DAY_MIN = 24 * 60.0;
        private const double DEFAULT_SPAN_MIN = 12 * 60.0;

        // model ve eksenler
        public PlotModel Model { get; }
        private readonly LinearAxis _x;       // dakika
        private readonly LinearAxis _y;       // km

        // WPF view referansı (hover/drag için)
        private WpfPlotView _view;

        // state
        private readonly Dictionary<string, List<LineSeries>> _move = new();
        private readonly Dictionary<string, List<LineSeries>> _dwell = new();
        private readonly Dictionary<string, OxyColor> _baseColor = new();
        private readonly SortedList<double, string> _stationsByKm = new();

        // sabit istasyon adları (sol tarafta metin)
        private readonly List<TextAnnotation> _stationLabels = new();
        // küçük saat etiketleri
        private readonly List<TextAnnotation> _timeLabels = new();

        // vurgular
        private HashSet<string> _fullRed = new();
        private HashSet<string> _unscheduled = new();
        private string _selectedTid;

        // hover/tooltip
        private readonly TextAnnotation _hover;
        private bool _hoverVisible;
        private string _hoverTid;

        // olaylar
        public Action<string> OnSeriesClicked;

        public bool ShowConflictOverlays { get; set; } = true;   // istersen kapat

        public PlotRenderer(IEnumerable<(string name, double km)> stations)
        {
            Model = new PlotModel
            {
                Background = OxyColors.White,
                TextColor = OxyColors.Black,
                PlotAreaBorderColor = OxyColors.Gray,
                PlotAreaBorderThickness = new OxyThickness(1)
            };

            // X: dakika
            _x = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Minimum = 0,
                Maximum = DEFAULT_SPAN_MIN,
                AbsoluteMinimum = -DAY_MIN,
                AbsoluteMaximum = 2 * DAY_MIN,
                Title = "Zaman",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot,
                MajorGridlineColor = OxyColor.FromAColor(40, OxyColors.Gray),
                MinorGridlineColor = OxyColor.FromAColor(20, OxyColors.Gray),
                IsPanEnabled = true,
                IsZoomEnabled = true,
                MajorStep = 60,     // 60 dk
                MinorStep = 15,
                LabelFormatter = (m) => ToHHMM(m)
            };
            _y = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "İstasyon",
                IsPanEnabled = true,
                IsZoomEnabled = true,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = OxyColor.FromAColor(28, OxyColors.Gray),
                MinorTickSize = 0
            };
            // ctor'da geçici formatter; ResetStations yeniden ayarlıyor
            _y.LabelFormatter = y =>
            {
                if (_stationsByKm.TryGetValue(y, out var name))
                    return $"{name} ({y:0.#} km)";
                return "";
            };

            Model.Axes.Add(_x);
            Model.Axes.Add(_y);

            // hover
            _hover = new TextAnnotation
            {
                Text = "",
                Stroke = OxyColors.Transparent,
                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                FontSize = 12,
                Background = OxyColor.FromAColor(230, OxyColors.LightYellow),
                TextColor = OxyColors.Black,
                Padding = new OxyThickness(6),
                Layer = AnnotationLayer.AboveSeries
            };
            Model.Annotations.Add(_hover);

            ResetStations(stations);
        }

        private void UpdateSmallLabelVisibility()
        {
            // Kullanıcı isteği: her zaman görünür olsun (performans pahasına)
            bool show = ALWAYS_SHOW_SMALL_TIME_LABELS ||
                        ((_x.ActualMaximum - _x.ActualMinimum) <= LABELS_DENSE_SPAN);

            var color = show ? OxyColors.Gray : OxyColors.Transparent;
            foreach (var ta in _timeLabels)
                ta.TextColor = color;

            Model.InvalidatePlot(false);
        }

        // === GENEL ===
        public void AttachTo(WpfPlotView view)
        {
            _view = view;

            var c = new PlotController();

            // Sol tık: en yakın çizgiyi seç
            c.BindMouseDown(OxyMouseButton.Left,
                new DelegatePlotCommand<OxyMouseDownEventArgs>((v, cc, a) =>
                {
                    HandleClick(a.Position);
                    a.Handled = true;
                }));

            // Orta tık veya Shift+Sol: pan
            c.BindMouseDown(OxyMouseButton.Middle, PlotCommands.PanAt);
            c.Bind(new OxyMouseDownGesture(OxyMouseButton.Left, OxyModifierKeys.Shift), PlotCommands.PanAt);

            // Tekerlek: zoom
            c.BindMouseWheel(PlotCommands.ZoomWheel);

            // Sağ tık: reset (12 saat)
            c.BindMouseDown(OxyMouseButton.Right,
                new DelegatePlotCommand<OxyMouseDownEventArgs>((v, cc, a) =>
                {
                    ResetView12h();
                    a.Handled = true;
                }));

            view.Controller = c;

            // WPF mouse hareketiyle kendi hover – OxyPlot sürümünden bağımsız
            view.MouseMove += (s, e) =>
            {
                var p = e.GetPosition(view);
                UpdateHover(new ScreenPoint(p.X, p.Y));
            };
            view.MouseLeave += (s, e) =>
            {
                HideHover();
            };
        }

        public void ResetStations(IEnumerable<(string name, double km)> stations)
        {
            var list = (stations ?? Array.Empty<(string, double)>()).OrderBy(s => s.km).ToList();
            if (list.Count == 0) list.Add(("?", 0));

            _stationsByKm.Clear();
            foreach (var s in list)
                if (!_stationsByKm.ContainsKey(s.km))
                    _stationsByKm.Add(s.km, s.name);

            var ymin = list.Min(s => s.km);
            var ymax = list.Max(s => s.km);
            if (ymax <= ymin) ymax = ymin + 1;

            // --- TÜM İSTASYONLAR GÖRÜNSÜN ---
            // 1) 1 km adım + üst/alt pad
            _y.MajorStep = 1.0;
            _y.MinorStep = 1.0;
            _y.Minimum = Math.Floor(ymin) - 0.5;
            _y.Maximum = Math.Ceiling(ymax) + 0.5;

            // 2) Yalnızca istasyon km'sine denk gelen tikte etiket yaz
            _y.LabelFormatter = (val) =>
            {
                foreach (var kv in _stationsByKm)
                {
                    if (Math.Abs(val - kv.Key) < 0.25) // küçük tolerans
                        return $"{kv.Value} ({kv.Key:0} km)";
                }
                return "";
            };

            // 3) Yatay çizgileri istasyonlara sabitle
            _y.ExtraGridlines = _stationsByKm.Keys.ToArray();
            _y.ExtraGridlineStyle = LineStyle.Solid;
            _y.ExtraGridlineColor = OxyColor.FromAColor(28, OxyColors.Gray);

            // Artık grafiğin içine istasyon adı yazmıyoruz
            foreach (var a in _stationLabels) Model.Annotations.Remove(a);
            _stationLabels.Clear();

            ClearAll();
            Model.InvalidatePlot(true);
        }

        private static double GuessNiceStep(List<(string name, double km)> st)
        {
            if (st.Count < 2) return 5.0;
            var diffs = st.Zip(st.Skip(1), (a, b) => Math.Abs(b.km - a.km)).Where(d => d > 1e-9);
            var minGap = diffs.Any() ? diffs.Min() : 5.0;
            double[] nice = { 0.5, 1, 2, 5, 10, 20, 50, 100 };
            foreach (var s in nice) if (s >= minGap - 1e-9) return s;
            return 5.0;
        }

        // === View kontrol (butonlar) ===
        public void ResetView12h()
        {
            _x.Zoom(0, DEFAULT_SPAN_MIN);   // görünümü doğrudan set et
            UpdateStationLabelX();
            UpdateSmallLabelVisibility();
            Model.InvalidatePlot(false);
        }

        public void PanX(double minutes)
        {
            var a = _x.ActualMinimum + minutes;
            var b = _x.ActualMaximum + minutes;
            _x.Zoom(a, b);
            UpdateStationLabelX();
            UpdateSmallLabelVisibility();
            Model.InvalidatePlot(false);
        }

        public void Zoom(double factor)
        {
            var mid = (_x.ActualMinimum + _x.ActualMaximum) / 2.0;
            var half = (_x.ActualMaximum - _x.ActualMinimum) * factor / 2.0;
            _x.Zoom(mid - half, mid + half);
            UpdateStationLabelX();
            UpdateSmallLabelVisibility();
            Model.InvalidatePlot(false);
        }

        public void SetX(double aMin, double bMin)
        {
            _x.Zoom(aMin, bMin);
            UpdateStationLabelX();
            UpdateSmallLabelVisibility();
            Model.InvalidatePlot(false);
        }

        public void PanY(double km)
        {
            _y.Zoom(_y.ActualMinimum + km, _y.ActualMaximum + km);
            Model.InvalidatePlot(false);
        }

        public void SetY(double aKm, double bKm)
        {
            _y.Minimum = aKm;
            _y.Maximum = bKm;
            Model.InvalidatePlot(false);
        }

        // === ÇİZİM ===
        public void ClearAll()
        {
            foreach (var t in _timeLabels) Model.Annotations.Remove(t);
            _timeLabels.Clear();

            Model.Series.Clear();
            _move.Clear();
            _dwell.Clear();
            _baseColor.Clear();

            HideHover();
            Model.InvalidatePlot(true);
        }

        public void DrawTrain(
            string tid,
            IList<(string name, double km)> pathStations,
            IDictionary<string, int> times,
            IDictionary<string, int> dwell,
            OxyColor baseColor,
            IList<((string u, string v) seg, int s, int e)> overlays,
            bool dashed = false, int? dashFromMin = null)
        {
            _baseColor[tid] = baseColor;
            _move[tid] = new List<LineSeries>();
            _dwell[tid] = new List<LineSeries>();
            dwell ??= new Dictionary<string, int>();
            overlays ??= Array.Empty<((string, string), int, int)>();

            for (int i = 0; i < pathStations.Count - 1; i++)
            {
                var u = pathStations[i];
                var v = pathStations[i + 1];
                if (!times.ContainsKey(u.name) || !times.ContainsKey(v.name))
                    continue;

                int arr = times[u.name];
                int dep = arr + (dwell.TryGetValue(u.name, out var d) ? d : 0);
                int arrV = times[v.name];

                // dwell (yatay)
                if (dep > arr)
                {
                    AddSegmentMaybeSplit(
                        tid, baseColor, dashed ? 1.2 : 1.4, dashed,
                        arr, u.km, dep, u.km, _dwell[tid], dashFromMin);
                    AddSmallTimeMark(arr, u.km);
                    AddSmallTimeMark(dep, u.km);
                }

                // hareket (eğimli)
                AddSegmentMaybeSplit(
                    tid, baseColor, dashed ? 1.2 : 1.6, dashed,
                    dep, u.km, arrV, v.km, _move[tid], dashFromMin);

                double y0 = u.km, y1 = v.km;
                if (Math.Abs(y1 - y0) > 1e-9)
                {
                    foreach (var yst in _stationsByKm.Keys)
                    {
                        if (yst < Math.Min(y0, y1) - 1e-9 || yst > Math.Max(y0, y1) + 1e-9) continue;
                        double r = (yst - y0) / (y1 - y0);
                        double tm = dep + (arrV - dep) * r;
                        AddSmallTimeMark(tm, yst);
                    }
                }

                if (ShowConflictOverlays)
                {
                    foreach (var ov in overlays)
                    {
                        if (Phys(ov.seg.u, ov.seg.v) != Phys(u.name, v.name)) continue;
                        double tv = arrV, td = dep; if (tv <= td) continue;
                        double ss = Math.Max(ov.s, td);
                        double ee = Math.Min(ov.e, tv);
                        if (ss < ee)
                        {
                            AddWrappedSegment(
                                tid: tid, color: OxyColors.Red, thick: 1.8, dashed: true,
                                x1Min: ss, y1: u.km, x2Min: ee, y2: v.km, bucket: null, tracker: true);
                        }
                    }
                }
            }

            RecolorAll();
            Model.InvalidatePlot(true);
            UpdateSmallLabelVisibility();
        }

        private void AddWrappedSegment(
            string tid, OxyColor color, double thick, bool dashed,
            double x1Min, double y1, double x2Min, double y2,
            List<LineSeries> bucket, bool tracker)
        {
            void AddOne(double a, double b, double c, double d)
            {
                var ls = new LineSeries
                {
                    Color = color,
                    StrokeThickness = thick,
                    LineStyle = dashed ? LineStyle.Dash : LineStyle.Solid,
                    Tag = tid,
                    Title = tid ?? "",
                    CanTrackerInterpolatePoints = true
                };

                ls.Points.Add(new DataPoint(a, b));
                ls.Points.Add(new DataPoint(c, d));
                Model.Series.Add(ls);
                if (bucket != null) bucket.Add(ls);
            }

            AddOne(x1Min, y1, x2Min, y2);
            AddOne(x1Min - DAY_MIN, y1, x2Min - DAY_MIN, y2);
            AddOne(x1Min + DAY_MIN, y1, x2Min + DAY_MIN, y2);
        }

        private void AddSegmentMaybeSplit(
            string tid, OxyColor color, double thick, bool baseDashed,
            double x1, double y1, double x2, double y2,
            List<LineSeries> bucket, int? dashFromMin)
        {
            if (dashFromMin == null)
            {
                AddWrappedSegment(tid, color, thick, baseDashed, x1, y1, x2, y2, bucket, tracker: true);
                return;
            }

            double t0 = dashFromMin.Value;

            if (x2 <= t0)
            {
                AddWrappedSegment(tid, color, thick, baseDashed, x1, y1, x2, y2, bucket, tracker: true);
                return;
            }
            if (x1 >= t0)
            {
                AddWrappedSegment(tid, color, thick, true, x1, y1, x2, y2, bucket, tracker: true);
                return;
            }

            double r = (t0 - x1) / (x2 - x1);
            double ym = y1 + r * (y2 - y1);

            AddWrappedSegment(tid, color, thick, baseDashed, x1, y1, t0, ym, bucket, tracker: true);
            AddWrappedSegment(tid, color, thick, true, t0, ym, x2, y2, bucket, tracker: true);
        }

        private void AddSmallTimeMark(double minutes, double km)
        {
            var ta = new TextAnnotation
            {
                Text = ToHHMM(minutes),
                TextPosition = new DataPoint(minutes, km),
                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                StrokeThickness = 0,
                FontSize = 10,
                TextColor = OxyColors.Gray,
                Layer = AnnotationLayer.AboveSeries
            };
            _timeLabels.Add(ta);
            Model.Annotations.Add(ta);
        }

        // === RENKLER / VURGULAR ===
        public void HighlightTrains(IEnumerable<string> ids)
        {
            _fullRed = new HashSet<string>(ids ?? Array.Empty<string>());
            RecolorAll();
        }

        public void MarkUnscheduled(IEnumerable<string> ids)
        {
            _unscheduled = new HashSet<string>(ids ?? Array.Empty<string>());

            foreach (var kv in _move)
                if (_unscheduled.Contains(kv.Key))
                    foreach (var ln in kv.Value)
                        ln.LineStyle = LineStyle.Dash;

            foreach (var kv in _dwell)
                if (_unscheduled.Contains(kv.Key))
                    foreach (var ln in kv.Value)
                        ln.LineStyle = LineStyle.Dash;

            Model.InvalidatePlot(false);
        }

        public void SetSelectedTrain(string tid)
        {
            _selectedTid = tid;
            RecolorAll();
        }

        private void RecolorAll()
        {
            foreach (var kv in _move)
            {
                var tid = kv.Key;
                var baseColor = _baseColor.TryGetValue(tid, out var c) ? c : OxyColors.SteelBlue;

                var color = _selectedTid == tid ? OxyColors.Orange
                           : _fullRed.Contains(tid) ? OxyColors.Red
                           : baseColor;

                var thick = _selectedTid == tid ? 2.6
                           : _fullRed.Contains(tid) ? 2.0
                           : 1.6;

                foreach (var ln in kv.Value)
                {
                    ln.Color = color;
                    ln.StrokeThickness = thick;
                }
            }

            foreach (var kv in _dwell)
            {
                var tid = kv.Key;
                var baseColor = _baseColor.TryGetValue(tid, out var c) ? c : OxyColors.SteelBlue;

                var color = _selectedTid == tid ? OxyColors.Orange
                           : _fullRed.Contains(tid) ? OxyColors.Red
                           : baseColor;

                var thick = _selectedTid == tid ? 2.2
                           : _fullRed.Contains(tid) ? 1.8
                           : 1.4;

                foreach (var ln in kv.Value)
                {
                    ln.Color = color;
                    ln.StrokeThickness = thick;
                }
            }

            Model.InvalidatePlot(false);
        }

        // === CLICK / HOVER ===
        private void HandleClick(ScreenPoint sp)
        {
            (LineSeries ls, TrackerHitResult hit, double d2) best = default;
            best.d2 = double.PositiveInfinity;

            foreach (var ls in Model.Series.OfType<LineSeries>())
            {
                if (ls.Tag == null) continue;

                var hit = ls.GetNearestPoint(sp, false);
                if (hit == null) continue;
                var d2 = (hit.Position.X - sp.X) * (hit.Position.X - sp.X) +
                         (hit.Position.Y - sp.Y) * (hit.Position.Y - sp.Y);
                if (d2 < best.d2)
                    best = (ls, hit, d2);
            }

            if (best.ls != null && best.d2 < 100)
            {
                var tid = best.ls.Tag as string;
                if (!string.IsNullOrWhiteSpace(tid))
                {
                    _selectedTid = tid;
                    RecolorAll();
                    OnSeriesClicked?.Invoke(tid);
                }
            }
        }

        private void UpdateHover(ScreenPoint sp)
        {
            if (_view == null) return;

            var now = _hoverWatch.ElapsedMilliseconds;
            if (now - _lastHoverMs < 30) return;
            _lastHoverMs = now;

            (LineSeries ls, TrackerHitResult hit, double d2) best = default;
            best.d2 = double.PositiveInfinity;

            foreach (var ls in Model.Series.OfType<LineSeries>())
            {
                if (ls.Tag == null) continue;

                var hit = ls.GetNearestPoint(sp, false);
                if (hit == null) continue;
                var d2 = (hit.Position.X - sp.X) * (hit.Position.X - sp.X) +
                         (hit.Position.Y - sp.Y) * (hit.Position.Y - sp.Y);
                if (d2 < best.d2)
                    best = (ls, hit, d2);
            }

            if (best.ls == null || best.d2 > 144)
            {
                HideHover();
                return;
            }

            var tid = best.ls.Tag as string ?? "";
            var dp = best.hit.DataPoint;
            _hover.Text = $"Tren: {tid}\nSaat: {ToHHMM(dp.X)}\nKm: {dp.Y:0.#}";
            _hover.TextPosition = new DataPoint(dp.X, dp.Y);
            _hoverTid = tid;
            _hoverVisible = true;
            Model.InvalidatePlot(false);
        }

        private void HideHover()
        {
            if (!_hoverVisible) return;
            _hover.Text = "";
            _hoverTid = null;
            _hoverVisible = false;
            Model.InvalidatePlot(false);
        }

        // === yardımcılar ===
        private static (string a, string b) Phys(string u, string v)
            => string.CompareOrdinal(u, v) < 0 ? (u, v) : (v, u);

        private static string ToHHMM(double minutes)
        {
            int m = (int)Math.Round(minutes);
            m = ((m % 1440) + 1440) % 1440;
            return $"{(m / 60) % 24:00}:{m % 60:00}";
        }

        private void AddStationNameAnnotations()
        {
            foreach (var a in _stationLabels) Model.Annotations.Remove(a);
            _stationLabels.Clear();

            double xLeft = _x.Minimum + 2.0;

            foreach (var (km, name) in _stationsByKm)
            {
                var ta = new TextAnnotation
                {
                    Text = name,
                    TextPosition = new DataPoint(xLeft, km),
                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                    StrokeThickness = 0,
                    TextColor = OxyColors.Black,
                    FontSize = 12,
                    Layer = AnnotationLayer.AboveSeries
                };
                _stationLabels.Add(ta);
                Model.Annotations.Add(ta);
            }
        }

        private void UpdateStationLabelX()
        {
            double xLeft = _x.ActualMinimum + 2.0;
            foreach (var ta in _stationLabels)
                ta.TextPosition = new DataPoint(xLeft, ta.TextPosition.Y);
        }
    }
}
