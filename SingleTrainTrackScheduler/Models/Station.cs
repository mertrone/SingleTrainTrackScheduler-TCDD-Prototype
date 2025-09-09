namespace SingleTrainTrackScheduler.Models
{
    public sealed class Station
    {
        public string Name { get; }
        public double Km { get; }
        public Station(string name, double km) { Name = name.Trim(); Km = km; }
        public override string ToString() => $"{Name} ({Km} km)";
    }
}
