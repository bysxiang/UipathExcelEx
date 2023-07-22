using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Models
{
    public class CustomKV<K, V>
    {
        public K Key { get; set; }

        public V Value { get; set; }

        public CustomKV()
        {
        }

        public CustomKV(K key, V value)
        {
            Key = key;
            Value = value;
        }

        public override bool Equals(object obj)
        {
            return obj is CustomKV<K, V> kV &&
                   EqualityComparer<K>.Default.Equals(Key, kV.Key) &&
                   EqualityComparer<V>.Default.Equals(Value, kV.Value);
        }

        public override int GetHashCode()
        {
            int hashCode = 206514262;
            hashCode = hashCode * -1521134295 + EqualityComparer<K>.Default.GetHashCode(Key);
            hashCode = hashCode * -1521134295 + EqualityComparer<V>.Default.GetHashCode(Value);
            return hashCode;
        }

        public static bool operator ==(CustomKV<K, V> left, CustomKV<K, V> right)
        {
            return EqualityComparer<CustomKV<K, V>>.Default.Equals(left, right);
        }

        public static bool operator !=(CustomKV<K, V> left, CustomKV<K, V> right)
        {
            return !(left == right);
        }
    }
}
