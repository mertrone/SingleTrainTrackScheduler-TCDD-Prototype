// Util/KeyValuePairDeconstruct.cs
namespace System.Collections.Generic
{
    // .NET Framework'te KeyValuePair için deconstruction ekler.
    public static class KeyValuePairExtensions
    {
        public static void Deconstruct<TKey, TValue>(
            this KeyValuePair<TKey, TValue> kvp,
            out TKey key,
            out TValue value)
        {
            key = kvp.Key;
            value = kvp.Value;
        }
    }
}
