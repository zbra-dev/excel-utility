using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelUtility.Utils
{
    public static class CollectionExtensions
    {
        private static Random random = new Random();

        public static void Sort<TSource>(this TSource[] source, Comparison<TSource> comparison)
        {
            Array.Sort<TSource>(source, comparison);
        }

        public static IEnumerable<TSource> OrderBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector, Comparison<TKey> comparison)
        {
            return source.OrderBy(keySelector, new DelegateComparer<TKey>(comparison));
        }

        public static IEnumerable<TSource> OrderBy<TSource>(this IEnumerable<TSource> source, Comparison<TSource> comparison)
        {
            return source.OrderBy(s => s, comparison);
        }

        public static IEnumerable<TSource> Distinct<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> knownKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (knownKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }

        public static IEnumerable<IEnumerable<T>> Split<T>(this IEnumerable<T> enumerable, int size)
        {
            List<T> list = new List<T>();
            foreach (T t in enumerable)
            {
                list.Add(t);
                if (list.Count == size)
                {
                    yield return list;
                    list.Clear();
                }
            }
            if (list.Count > 0)
                yield return list;
        }

        public static V GetOrDefault<K, V>(this IDictionary<K, V> dict, K key)
        {
            V v;
            if (dict.TryGetValue(key, out v))
                return v;
            return default(V);
        }

        public static void AddRange<T>(this ICollection<T> collection, IEnumerable<T> enumerable)
        {
            foreach (T t in enumerable)
                collection.Add(t);
        }

        public static HashSet<T> ToHashSet<T>(this IEnumerable<T> enumerable)
        {
            return new HashSet<T>(enumerable);
        }

        public static T RandomValue<T>(this IEnumerable<T> enumerable)
        {
            T[] array = enumerable.ToArray();
            if (array.Length == 0)
                return default(T);
            if (array.Length == 1)
                return array[0];
            return array[random.Next(array.Length)];
        }

        public static int BinarySearch<T>(this T[] array, T value, Comparison<T> comparison)
        {
            return Array.BinarySearch<T>(array, value, new DelegateComparer<T>(comparison));
        }

        public static int BinarySearch<T>(this List<T> list, T value, Comparison<T> comparison)
        {
            return list.BinarySearch(value, new DelegateComparer<T>(comparison));
        }

        public static IMultiMap<K, V> ToMultiMap<K, V>(this IEnumerable<V> col, Func<V, K> keySelector)
        {
            MultiMap<K, V> map = new MultiMap<K, V>();
            foreach (V v in col)
                map.Add(keySelector(v), v);
            return map;
        }
    }
}
