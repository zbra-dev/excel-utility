using System;
using System.Collections.Generic;

namespace ExcelUtility.Utils
{
    [Serializable]
    public class MultiMap<K, V> : IMultiMap<K, V>
    {
        public delegate ICollection<V> CollectionFactory();

        private Dictionary<K, ICollection<V>> map;
        private CollectionFactory collectionFactory;
        private bool autoClean;

        public MultiMap(CollectionFactory collectionFactory, bool autoClean)
        {
            if (collectionFactory == null)
                throw new ArgumentException("collection factory cannot be null");
            this.map = new Dictionary<K, ICollection<V>>();
            this.collectionFactory = collectionFactory;
            this.autoClean = autoClean;
        }

        public MultiMap(CollectionFactory collectionFactory)
            : this(collectionFactory, true)
        {
        }

        public MultiMap() 
            : this(() => new List<V>(), true)
        {
        }

        public MultiMap(IMultiMap<K, V> map)
        {
            this.map = new Dictionary<K, ICollection<V>>(map.Count);
            this.collectionFactory = () => new List<V>();
            this.autoClean = true;
            AddRange(map);
        }

        public void AddRange(IMultiMap<K, V> map)
        {
            foreach (K k in map.Keys)
                foreach (V v in map[k])
                    Add(k, v);
        }

        public void Add(K key, V value)
        {
            EnsureCollection(key).Add(value);
        }

        public ICollection<V> EnsureCollection(K key)
        {
            ICollection<V> list = null;
            if (!map.TryGetValue(key, out list))
            {
                list = collectionFactory();
                map.Add(key, list);
            }
            return list;
        }

        public void Clear()
        {
            map.Clear();
        }

        public bool ContainsKey(K key)
        {
            return map.ContainsKey(key);
        }

        public bool Remove(K key)
        {
            return map.Remove(key);
        }

        public bool RemoveAllValues(K key)
        {
            ICollection<V> list = null;
            if (map.TryGetValue(key, out list))
            {
                bool changed = list.Count > 0;
                if (autoClean)
                    map.Remove(key);
                else if (changed)
                    list.Clear();
                return changed;
            }
            return false;
        }

        public bool RemoveValue(K key, V value)
        {
            ICollection<V> list = null;
            if (map.TryGetValue(key, out list))
            {
                bool changed = list.Remove(value);
                if (autoClean && list.Count == 0)
                    map.Remove(key);
                return changed;
            }
            return false;
        }

        public bool ContainsValue(V value)
        {
            foreach (ICollection<V> list in map.Values)
            {
                if (list.Contains(value))
                    return true;
            }
            return false;
        }

        public bool ContainsValue(K key, V value)
        {
            ICollection<V> list = null;
            if (map.TryGetValue(key, out list))
                return list.Contains(value);
            return false;
        }

        public bool TryGetValue(K key, out ICollection<V> value)
        {
            return map.TryGetValue(key, out value);
        }

        public ICollection<K> Keys { get { return map.Keys; } }

        public ICollection<ICollection<V>> Values { get { return map.Values; } }

        public ICollection<V> this[K key] { get { return map[key]; } }

        public int Count { get { return map.Count; } }

        public bool IsReadOnly { get { return false; } }

        public IEnumerator<V> GetEnumerator()
        {
            foreach (ICollection<V> list in map.Values)
                foreach (V v in list)
                    yield return v;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
