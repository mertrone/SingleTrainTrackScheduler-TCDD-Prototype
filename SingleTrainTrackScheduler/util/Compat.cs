using System.Collections.Generic;

namespace SingleTrainTrackScheduler.util
{
    // .NET 6 API'lerini .NET 4.8'de taklit
    internal static class Compat
    {
        public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> d, TKey key, TValue value)
        {
            if (d.ContainsKey(key)) return false;
            d.Add(key, value);
            return true;
        }

        public static TValue GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> d, TKey key, TValue defaultValue = default(TValue))
        {
            TValue v;
            return d != null && d.TryGetValue(key, out v) ? v : defaultValue;
        }

        // Math.Clamp muadili
        public static int Clamp(int value, int min, int max)
            => value < min ? min : (value > max ? max : value);

        public static double Clamp(double value, double min, double max)
            => value < min ? min : (value > max ? max : value);
    }
}
