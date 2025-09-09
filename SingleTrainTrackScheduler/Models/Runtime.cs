namespace SingleTrainTrackScheduler.Models
{
    public class Runtime
    {
        public string U { get; set; }
        public string V { get; set; }
        public string Type { get; set; }    // P / F / G
        public int RunMin { get; set; }     // dakika

        // Eski kodla uyumluluk (Engine'de TType yazılı yerlere takılmaması için)
        public string TType
        {
            get => Type;
            set => Type = value;
        }

        public Runtime() { }

        public Runtime(string u, string v, string type, int runMin)
        {
            U = u; V = v; Type = type; RunMin = runMin;
        }
    }
}

