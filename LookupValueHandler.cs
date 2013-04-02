using System.Linq;

namespace Migrate
{
    abstract class LookupValueHandler<T> : ILookupValueHandler where T : class, new()
    {
        public object Create(int id)
        {
            dynamic fieldValue = new T();
            fieldValue.LookupId = id;
            return fieldValue;
        }

        protected abstract int GetId(T lookupField);

        public int[] Extract(object lookupField) 
        {
            var single = lookupField as T;
            if (single != null)
                return new[] { GetId(single) };

            var multiple = lookupField as T[];
            return multiple == null ? new int[0] : multiple.Select(GetId).ToArray();
        }
    }
}
