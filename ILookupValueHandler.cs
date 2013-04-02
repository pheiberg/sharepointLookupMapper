namespace Migrate
{
    internal interface ILookupValueHandler
    {
        object Create(int id);
        int[] Extract(object lookupField);
    }
}