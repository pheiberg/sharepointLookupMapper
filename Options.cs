using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;

namespace Migrate
{
    public class Options
    {
        [Option('u', "user", HelpText = "The user name. If not provided, use the current user.")]
        public string User { get; set; }

        [Option('p', "password", HelpText = "The password for the user.")]
        public string Password { get; set; }

        [Option('m', "master", HelpText = "The name of the master list to set the lookups in.", Required = true)]
        public string Master { get; set; }
        
        [Option('l', "lookup", HelpText = "The name of the lookup column.", Required = true)]
        public string Lookup { get; set; }

        [Option('s', "simulate", HelpText = "Only simulates the operation and prints out what would have been performed instead of actually changing the lookup values.")]
        public bool Simulate { get; set; }

        [OptionList('i', "identifiers", HelpText = "(Default: Title). The columns that can uniquely identify an item in the list (instead of Id).", DefaultValue = new []{ "Title" }, Separator = ',')]
        public IList<string> IdentifyingColumns { get; set; }
        
        [OptionList("lookup-identifiers", HelpText = "(Default: Title). The columns that can uniquely identify an item in the lookup list (instead of Id).", DefaultValue = new []{ "Title" }, Separator = ',')]
        public IList<string> IdentifyingLookupColumns { get; set; }

        [Option("page-size", DefaultValue = 5000, HelpText = "Number of items to fetch from main list at a time")]
        public int PageSize { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
        }


    }
}
