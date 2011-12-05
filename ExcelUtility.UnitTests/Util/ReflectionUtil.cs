using System;
using System.Linq;
using System.Reflection;

namespace ExcelUtility.UnitTests.Util
{
    public class ReflectionUtil
    {
        public object GetValue(object target, string name)
        {
            return TryGetFieldInfo(target, target.GetType(), name);
        }

        private object TryGetPropertyInfo(object target, Type type, string name)
        {
            var propertyInfos = type.GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
            var property = propertyInfos.Where(p => p.Name == name).FirstOrDefault();
            //return property != null ? property.GetValue(target) : null;
            return null;
        }

        public object TryGetFieldInfo(object target, Type type, string fieldName)
        {
            var fields = type.GetFields(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
            var field = fields.Where(f => f.Name == fieldName).FirstOrDefault();
            return field != null ? field.GetValue(target) : null;
        }
    }
}
