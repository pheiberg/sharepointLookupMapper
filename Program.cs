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

            var sourceLookupList = GetList(sourceContext, sourceWeb, new Guid(((FieldLookup)(sourceLookup)).LookupList));
            Console.WriteLine("Loading source lookup items from {0} ...", sourceLookupList.Title);
            var sourceLookupItems = GetAllItems(sourceContext, sourceLookupList);

            var destinationLookupList = GetList(destinationContext, destinationWeb, new Guid(((FieldLookup)(destinationLookup)).LookupList));
            Console.WriteLine("Loading destination lookup items from {0} ...", destinationLookupList.Title);
            var destinationLookupItems = GetAllItems(destinationContext, destinationLookupList);

            Console.WriteLine("Mapping lookup tables ...");
            IDictionary<int, ListMappings> lookupMappings = GetLookupMappings(sourceLookupItems, destinationLookupItems, options.IdentifyingLookupColumns);
            
            Console.WriteLine("Checking for duplicates ...");
            var lookupDuplicates = lookupMappings.GroupBy(i => i.Value.SourceId).Where(g => g.Count() > 1).ToList();
            foreach (var duplicate in lookupDuplicates)
            {
                Console.WriteLine("Lookup duplicate: " + duplicate.Key + " => " + string.Join(", ", duplicate.Select(d => d.Value.DestinationId)));
            }
            if (lookupDuplicates.Any())
            {
                Environment.Exit(1);
            }

            Console.WriteLine("Loading source items ...");
            var sourceItems = GetAllItems(sourceContext, sourceList);

            Console.WriteLine("Loading destination items ...");
            var destinationItems = GetAllItems(destinationContext, destinationList);

            Console.WriteLine("Mapping items ...");
            IList<MasterItemMapping> itemMappings = GetItemMappings(sourceLookup, lookupMappings, destinationItems, sourceItems, options.IdentifyingColumns);

            Console.WriteLine("Checking for duplicates ...");
            var duplicates = itemMappings.GroupBy(i => i.SourceId).Where(g => g.Count() > 1).ToList();
            foreach (var duplicate in duplicates)
            {
                Console.WriteLine("Master duplicate: " + duplicate.Key + " => " + string.Join(", ", duplicate.Select(d => d.DestinationId)));
            }
            if(duplicates.Any())
            {
                Environment.Exit(1);
            }

            if(options.Simulate)
            {
                PrintMappings(itemMappings);
            }
            else
            {
                UpdateMappingsAtDestination(itemMappings.ToDictionary(i => i.DestinationId, i => i), destinationItems, options.Lookup, (FieldLookup)destinationLookup);
            }
        }

        private static IList<MasterItemMapping> GetItemMappings(Field sourceLookup, IDictionary<int, ListMappings> lookupMappings, IEnumerable<ListItem> destinationItems, IEnumerable<ListItem> sourceItems, IEnumerable<string> identifyingColumns)
        {
            Func<ListItem, ListItem, bool> isEqual = (source, destination) => identifyingColumns.All(column => source[column].Equals(destination[column]));
            return (from sourceItem in sourceItems
                    from destinationItem in destinationItems
                    where isEqual(sourceItem, destinationItem)
                    let lookups = ExtractLookupIds(sourceItem[sourceLookup.InternalName])
                    select
                        new MasterItemMapping
                            {
                                SourceId = sourceItem.Id,
                                SourceTitle = (string)sourceItem["Title"],
                                SourceLookupIds = lookups,
                                DestinationId = destinationItem.Id,
                                DestinationTitle = (string)destinationItem["Title"],
                                DestinationLookupIds  = GetCorrespondingLookups(lookupMappings, lookups).ToArray()
                            }).ToList();
        }

        private static Dictionary<int, ListMappings> GetLookupMappings(IEnumerable<ListItem> sourceLookupItems, IEnumerable<ListItem> destinationLookupItems, IEnumerable<string> identifyingColumns)
        {
            Func<ListItem, ListItem, bool> isEqual = (source, destination) => identifyingColumns.All(column => source[column].Equals(destination[column]));
            return (from sourceLookupItem in sourceLookupItems
                    from destinationLookupItem in destinationLookupItems
                    where isEqual(sourceLookupItem, destinationLookupItem)
                    select new ListMappings
                               {
                                   SourceId = sourceLookupItem.Id,
                                   SourceTitle = (string)sourceLookupItem["Title"],
                                   DestinationId = destinationLookupItem.Id,
                                   DestinationTitle = (string)destinationLookupItem["Title"],
                               }).ToDictionary(item => item.SourceId, item => item);
        }

        private static void PrintMappings(IEnumerable<dynamic> itemMappings)
        {
            Console.WriteLine("Source\t\tDestination");
            foreach (var itemMapping in itemMappings)
            {
                Console.WriteLine("s:{0}-\"{1}\" sl:{2} \t d:{3}-\"{4}\" dl:{5}", itemMapping.SourceId, itemMapping.SourceTitle,
                                  string.Join<int>(", ", itemMapping.SourceLookupIds), itemMapping.DestinationId,
                                  itemMapping.DestinationTitle, string.Join<int>(", ", itemMapping.DestinationLookupIds));
            }
        }

        private static void UpdateMappingsAtDestination(Dictionary<int, MasterItemMapping> itemMappings, IEnumerable<ListItem> destinationItems, string lookup, FieldLookup destinationLookup)
        {
            Console.WriteLine("Updating lookup values ...");

            foreach (var destinationItem in destinationItems.Where(item => itemMappings.ContainsKey(item.Id)))
            {
                var mapping = itemMappings[destinationItem.Id];

                if(!mapping.DestinationLookupIds.Any() || !mapping.SourceLookupIds.Any())
                    continue;

                object value = destinationLookup.AllowMultipleValues ? (object) mapping.DestinationLookupIds.Select( id => new FieldLookupValue { LookupId = id} )
                                   : new FieldLookupValue { LookupId = mapping.DestinationLookupIds.First() };
                destinationItem[lookup] = value;
                destinationItem.Update();
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

        private static IList<ListItem> GetAllItems(ClientRuntimeContext context, List list)
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
}