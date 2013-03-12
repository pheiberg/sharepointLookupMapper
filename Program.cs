using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using CommandLine;
using Microsoft.SharePoint.Client;

namespace Migrate
{
    class Program
    {
        private static NetworkCredential _networkCredential;
        private const string SourceSiteUrl = "http://lsc.alfalaval.org/sites/c_ciis/c_one/One4AL%20Test/";
        private const string TargetSiteUrl = "http://work.alfalaval.org/temporary/one4al";

        static void Main(string[] args)
        {
            var options = new Options();
            var parser = new Parser(configuration =>
                                        {
                                            configuration.IgnoreUnknownArguments = false;
                                            configuration.HelpWriter = Console.Error;
                                        });
            parser.ParseArgumentsStrict(args, options);

            _networkCredential = string.IsNullOrEmpty(options.User) ? CredentialCache.DefaultNetworkCredentials : new NetworkCredential(options.User, options.Password);

            var sourceContext = new ClientContext(SourceSiteUrl) {Credentials = _networkCredential};
            var sourceWeb = GetWeb(sourceContext);

            Console.WriteLine("Loading source list {0} ...", options.Master);
            var sourceList = GetList(sourceContext, sourceWeb, options.Master);

            var sourceLookup = GetLookupField(sourceContext, sourceList, options.Lookup);
            ValidateLookupField(sourceLookup, "source", options.Lookup);
            
            var destinationContext = new ClientContext(TargetSiteUrl) { Credentials = _networkCredential };
            var destinationWeb = GetWeb(destinationContext);
            Console.WriteLine("Loading destination list {0} ...", options.Master);
            var destinationList = GetList(destinationContext, destinationWeb, options.Master);

            var destinationLookup = GetLookupField(destinationContext, destinationList, options.Lookup);
            ValidateLookupField(destinationLookup, "destination", options.Lookup);

            Console.WriteLine("Loading source lookup items ...");
            var sourceLookupList = GetList(sourceContext, sourceWeb, new Guid(((FieldLookup)(sourceLookup)).LookupList));
            var sourceLookupItems = GetAllItems(sourceContext, sourceLookupList);

            Console.WriteLine("Loading destination lookup items ...");
            var destinationLookupList = GetList(destinationContext, destinationWeb, new Guid(((FieldLookup)(destinationLookup)).LookupList));
            var destinationLookupItems = GetAllItems(destinationContext, destinationLookupList);

            Console.WriteLine("Mapping lookup tables ...");
            var lookupMappings = (from sourceLookupItem in sourceLookupItems
                                 from destinationLookupItem in destinationLookupItems
                                 where (string)sourceLookupItem["Title"] == (string)destinationLookupItem["Title"]
                                 select new ListMappings
                                 {
                                     SourceId = sourceLookupItem.Id,
                                     SourceTitle = (string)sourceLookupItem["Title"],
                                     DestinationId = destinationLookupItem.Id,
                                     DestinationTitle = (string)destinationLookupItem["Title"],
                                 }).ToDictionary(item => item.SourceId, item => item);

            Console.WriteLine("Loading source items ...");
            var sourceItems = GetAllItems(sourceContext, sourceList);

            Console.WriteLine("Loading destination items ...");
            var destinationItems = GetAllItems(destinationContext, destinationList);

            Console.WriteLine("Mapping items ...");
            var itemMappings = from sourceItem in sourceItems
                               from destinationItem in destinationItems
                               where (string)sourceItem["Title"] == (string)destinationItem["Title"]
                               let lookups = ExtractLookupIds(sourceItem[sourceLookup.InternalName])
                               select
                                   new
                                       {
                                           SourceId = sourceItem.Id,
                                           SourceTitle = (string)sourceItem["Title"],
                                           SourceLookupIds = lookups,
                                           DestinationId = destinationItem.Id,
                                           DestinationTitle = (string)destinationItem["Title"],
                                           DestinationLookupIds  = GetCorrespondingLookups(lookupMappings, lookups).ToArray()
                                       };

            
            Console.WriteLine("Source\t\tDestination");
            foreach (var itemMapping in itemMappings)
            {
                Console.WriteLine("s:{0}-\"{1}\" sl:{2} \t d:{3}-\"{4}\" dl:{5}", itemMapping.SourceId, itemMapping.SourceTitle, string.Join(", ", itemMapping.SourceLookupIds), itemMapping.DestinationId, itemMapping.DestinationTitle, string.Join(", ", itemMapping.DestinationLookupIds));
            }
        }

        private static IEnumerable<int> GetCorrespondingLookups(IDictionary<int, ListMappings> lookupMappings, IEnumerable<int> lookups)
        {
            foreach (var lookup in lookups)
            {
                ListMappings mapping;
                if (lookupMappings.TryGetValue(lookup, out mapping))
                    yield return mapping.DestinationId;
            }
        }

        private static int[] ExtractLookupIds(object lookupField)
        {
            var single = lookupField as FieldLookupValue;
            if(single != null)
                return new[]{ single.LookupId};

            var multiple = lookupField as FieldLookupValue[];
            return multiple == null ? new int[0] : multiple.Select(l => l.LookupId).ToArray();
        }

        private static Web GetWeb(ClientContext context)
        {
            var web = context.Web;
            context.Load(web, w => w.Lists);
            context.ExecuteQuery();
            return web;
        }

        private static List GetList(ClientRuntimeContext context, Web web, string listTitle)
        {
            var sourceList = web.Lists.GetByTitle(listTitle);
            context.Load(sourceList);
            context.ExecuteQuery();
            return sourceList;
        }
        
        private static List GetList(ClientRuntimeContext context, Web web, Guid listId)
        {
            var sourceList = web.Lists.GetById(listId);
            context.Load(sourceList);
            context.ExecuteQuery();
            return sourceList;
        }

        private static IEnumerable<ListItem> GetAllItems(ClientRuntimeContext context, List list)
        {
            var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
            context.Load(items);
            context.ExecuteQuery();
            return items.ToList();
        }

        private static void ValidateLookupField(Field lookupField, string source, string lookupName)
        {
            if (lookupField == null)
            {
                Console.Error.WriteLine("Field \"{0}\" in {1} cannot be found", lookupName, source);
                Environment.Exit(0);
            }

            if (lookupField.FieldTypeKind != FieldType.Lookup)
            {
                Console.Error.WriteLine("Field \"{0}\" in {1} is not a lookup", lookupName, source);
                Environment.Exit(0);
            }
        }

        private static Field GetLookupField(ClientRuntimeContext sourceContext, List sourceList, string lookupName)
        {
            var sourceFields = sourceList.Fields;
            sourceContext.Load(sourceFields);
            sourceContext.ExecuteQuery();

            return sourceFields.ToList().SingleOrDefault(f => f.Title == lookupName);
        }
    }

    public class ListMappings
    {
        public int SourceId { get; set; }
        public string SourceTitle { get; set; }
        public int DestinationId { get; set; }
        public string DestinationTitle { get; set; }
    }
}