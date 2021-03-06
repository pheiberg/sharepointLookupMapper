﻿using System;
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
        const int BatchSize = 200;
        const int MaxListPageSize = 5000;
        private static readonly LookupValueHandlerFactory LookupValueHandlerFactory = new LookupValueHandlerFactory();

        static void Main(string[] args)
        {
            var options = GetOptions(args);

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
            var sourceLookupItems = GetAllItems(sourceContext, sourceLookupList, options.PageSize);

            var destinationLookupList = GetList(destinationContext, destinationWeb, new Guid(((FieldLookup)(destinationLookup)).LookupList));
            Console.WriteLine("Loading destination lookup items from {0} ...", destinationLookupList.Title);
            var destinationLookupItems = GetAllItems(destinationContext, destinationLookupList, options.PageSize);

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
            var sourceItems = GetAllItems(sourceContext, sourceList, options.PageSize);

            Console.WriteLine("Loading destination items ...");
            var destinationItems = GetAllItems(destinationContext, destinationList, options.PageSize);

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
                UpdateMappingsAtDestination(itemMappings.ToDictionary(i => i.DestinationId, i => i), destinationItems, (FieldLookup)destinationLookup, destinationContext);
            }
        }

        private static Options GetOptions(string[] args)
        {
            var options = new Options();
            var parser = new Parser(configuration =>
                                        {
                                            configuration.IgnoreUnknownArguments = false;
                                            configuration.HelpWriter = Console.Error;
                                        });
            parser.ParseArgumentsStrict(args, options);
            if (options.PageSize > MaxListPageSize || options.PageSize < 1)
            {
                options.PageSize = MaxListPageSize;
            }
            return options;
        }

        private static IList<MasterItemMapping> GetItemMappings(Field sourceLookup, IDictionary<int, ListMappings> lookupMappings, IEnumerable<ListItem> destinationItems, IEnumerable<ListItem> sourceItems, IEnumerable<string> identifyingColumns)
        {
            var lookupValueHandler = LookupValueHandlerFactory.Create(sourceLookup.FieldTypeKind);
            Func<ListItem, ListItem, bool> isEqual = (source, destination) => identifyingColumns.All(column => 
                "Id".Equals(column, StringComparison.InvariantCultureIgnoreCase) ? source.Id == destination.Id : source[column].Equals(destination[column]));
            return (from sourceItem in sourceItems
                    from destinationItem in destinationItems
                    where isEqual(sourceItem, destinationItem)
                    let lookups = lookupValueHandler.Extract(sourceItem[sourceLookup.InternalName])
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

        private static void UpdateMappingsAtDestination(Dictionary<int, MasterItemMapping> itemMappings, IEnumerable<ListItem> destinationItems, FieldLookup destinationLookup, ClientContext destinationContext)
        {
            Console.WriteLine("Updating lookup values ...");

            int batchedItems = 0;
            foreach (var destinationItem in destinationItems.Where(item => itemMappings.ContainsKey(item.Id)))
            {
                var mapping = itemMappings[destinationItem.Id];

                if(!mapping.DestinationLookupIds.Any() || !mapping.SourceLookupIds.Any())
                    continue;

                var creator = LookupValueHandlerFactory.Create(destinationLookup.FieldTypeKind);
                object value = destinationLookup.AllowMultipleValues ? mapping.DestinationLookupIds.Select(creator.Create).ToArray()
                                   : creator.Create(mapping.DestinationLookupIds.First());
                destinationItem[destinationLookup.InternalName] = value;
                destinationItem.Update();
                batchedItems++;
                
                if(batchedItems % BatchSize == 0)
                {
                    destinationContext.ExecuteQuery();
                    batchedItems = 0;
                }
            }
            
            if(batchedItems > 0)
            {
                destinationContext.ExecuteQuery();
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

        private static IList<ListItem> GetAllItems(ClientRuntimeContext context, List list, int pageSize = MaxListPageSize)
        {
            ListItemCollectionPosition position = null;
            IEnumerable<ListItem> results = Enumerable.Empty<ListItem>();
            
            do
            {
                var query = new CamlQuery
                {
                    ListItemCollectionPosition = position,
                    ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query></Query><RowLimit>{0}</RowLimit></View>", pageSize)
                };

                var items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                position = items.ListItemCollectionPosition;
                results = results.Concat(items);
            } 
            while (position != null);
            
            return results.ToList();
        }

        private static void ValidateLookupField(Field lookupField, string source, string lookupName)
        {
            if (lookupField == null)
            {
                Console.Error.WriteLine("Field \"{0}\" in {1} cannot be found", lookupName, source);
                Environment.Exit(0);
            }

            if (lookupField.FieldTypeKind != FieldType.Lookup && lookupField.FieldTypeKind != FieldType.User)
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