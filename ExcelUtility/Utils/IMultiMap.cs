using System.Collections.Generic;

namespace ExcelUtility.Utils
{
    public interface IMultiMap<K, V> : IEnumerable<V>
    {
        void AddRange(IMultiMap<K, V> map);
        void Add(K key, V value);
        void Clear();
        bool ContainsKey(K key);
        bool Remove(K key);
        bool RemoveValue(K key, V value);
        bool RemoveAllValues(K key);
        bool ContainsValue(V value);
        bool ContainsValue(K key, V value);
        bool TryGetValue(K key, out ICollection<V> value);
        ICollection<K> Keys { get; }
        ICollection<ICollection<V>> Values { get; }
        ICollection<V> this[K key] { get; }
        int Count { get; }
        bool IsReadOnly { get; }
        ICollection<V> EnsureCollection(K key);
    }
}
