using CommandLine;

namespace SlideTester.App.CommandLine
{
    /// <summary>
    /// This class is the command line definition for the Manifest verb.
    /// This verb handles processing of "debug" task, reading serialized task manifest files. 
    /// </summary>
    [Verb("uri", HelpText = "Processes ppt from s3, unc or local path uri, writes output results to folder")]
    public class UriOptions : BaseOptions
    {
        [Value(
            0, // Start at index = 0 when populating
            MetaName = "uri",
            HelpText = "Uri to local, UNC or s3 powerpoint file to process",
            Required = true)]
        public string PowerpointFileUri { get; set; } = string.Empty;
        
        [Option(
            'r',
            "region",
            Required = false,
            HelpText = "aws s3 regional endpoint (e.g. us-east-1), only required if uri is to an s3 resource")]
        public string Region { get; set; } = string.Empty;
    }
}
