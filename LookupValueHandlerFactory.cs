using Microsoft.SharePoint.Client;

namespace Migrate
{
    class LookupValueHandlerFactory
    {
        public ILookupValueHandler Create(FieldType fieldType)
        {
            return fieldType == FieldType.User ? (ILookupValueHandler) new UserLookupValueHandler() : new FieldLookupValueHandler();
        }
    }
}
