using Microsoft.SharePoint.Client;

namespace Migrate
{
    class FieldLookupValueHandler : LookupValueHandler<FieldLookupValue>
    {
        protected override int GetId(FieldLookupValue lookupField)
        {
            return lookupField.LookupId;
        }
    }
}