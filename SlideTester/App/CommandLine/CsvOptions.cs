using CommandLine;

namespace SlideTester.App.CommandLine
{
    [Verb("csv", HelpText = "Reads a csv containing a list of PPT uris. Processes each uri and writes output")]
    public class CsvOptions : BaseOptions
    {
        [Value(
            0, // Start at index = 0 when populating
            MetaName = "csv",
            HelpText = "Path to .csv file containing list of local, UNC or s3 powerpoint files to process",
            Required = true)]
        public string CsvFilePath { get; set; } = string.Empty;


        [Option(
            "hasHeaderRecord",
            Required = false,
            Default = false,
            HelpText = "If set then csv parsing will assume that a header record exists in the first "
                       + "csv row. Else parsing will assume the first csv row maps to a record to process")]
        public bool HasHeaderRecord { get; set; } = false;
    }
}
