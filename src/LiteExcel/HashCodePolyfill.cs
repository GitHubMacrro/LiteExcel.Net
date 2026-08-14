#if NET48
namespace System
{
    // net48 doesn't have System.HashCode, provide a simple polyfill
    internal sealed class HashCode
    {
        private int _hash = 17;
        public void Add<T>(T value) { _hash = _hash * 31 + (value?.GetHashCode() ?? 0); }
        public int ToHashCode() => _hash;
    }
}
#endif
