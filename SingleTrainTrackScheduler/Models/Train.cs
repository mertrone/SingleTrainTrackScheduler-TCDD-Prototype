using System.Collections.Generic;

namespace SingleTrainTrackScheduler.Models
{
    public sealed class Train
    {
        public string Id { get; }
        public string Type { get; set; }          // 'P','F','G'
        public string Direction { get; set; }     // 'up'/'down'
        public string Origin { get; set; }
        public int ReqTime { get; set; }          // 0..1440
        public string Dest { get; set; } = "";

        // station -> dwell minutes
        public Dictionary<string, int> Dwell { get; } = new();

        public Train(string id, string type, string dir, string origin, int reqTime, string dest = "")
        {
            Id = id.Trim();
            Type = (type ?? "P").Trim().ToUpper();
            Direction = (dir ?? "up").Trim().ToLower();
            Origin = origin.Trim();
            ReqTime = reqTime;
            Dest = dest?.Trim() ?? "";
        }
        public override string ToString() => $"{Id} [{Type},{Direction}] {Origin}->{(string.IsNullOrWhiteSpace(Dest) ? "END" : Dest)} @{ReqTime}m";
    }
}