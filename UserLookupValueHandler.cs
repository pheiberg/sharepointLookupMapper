using Microsoft.SharePoint.Client;

namespace Migrate
{
    class UserLookupValueHandler : LookupValueHandler<FieldUserValue>
    {
        protected override int GetId(FieldUserValue lookupField)
        {
            return lookupField.LookupId;
        }
    }
}