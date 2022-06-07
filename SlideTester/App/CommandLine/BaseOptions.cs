using CommandLine;
using Serilog.Events;

namespace SlideTester.App.CommandLine
{
    /// <summary>
    /// Abstract class to configure common commandline option between all SlideTester verbs
    /// </summary>
    public abstract class BaseOptions
    {
        [Option(
            'o',
            "outputPath",
            Required = true,
            HelpText = "Folder to which all results will be written")]
        public string OutputPath { get; set; } = string.Empty;
        
        [Option(
            "compareOutput",
            Required = false,
            Default = false,
            HelpText = "If set then out of each slide driver for each input file will be compared and written to compareResult.json")]
        public bool ShouldCompareOutput { get; set; } = false;

        [Option(
            "printCompareResults",
            Required = false,
            Default = false,
            HelpText = "If set and compareOutput is set then we will write compare results to console")]
        public bool ShouldPrintCompareResults { get; set; } = false;

        [Option(
            "skipImageCompare",
            Required = false,
            Default = false,
            HelpText = "If set and compareOutput is set then we will only compare slide metadata (text) output")]
        public bool SkipImageCompare { get; set; } = false;
        
        [Option(
            'p',
            "maxSlideParallelization",
            Required = false,
            Default = 10,
            HelpText = "Max number of slide to concurrently process _per_ slide deck")]
        public int MaxSlideParallelization { get; set; } = 10;
        
        [Option(
            'e',
            "eventLog", 
            Required = false, 
            Default =  LogEventLevel.Information,
            HelpText = "Sets process to use a console log sink for event logging (only available on Windows).\n"
             + "Parameter to this option must be valid LogLevel values = (Verbose,Debug,Information,Warning,Error,Fatal)\n"
             + "Note: consoleLog, eventLog and fileLog options can be combined.\n"
             + "Note 2: If neither consoleLog, eventLog nor fileLog options are supplied the default (in settings) logging configuration is used.\n"
             + "Example: '--eventLog warning'")]
        public LogEventLevel EventLogConfig { get; set; } = LogEventLevel.Information;

        [Option(
            'c',
            "consoleLog",
            Required = false,
            Default = LogEventLevel.Information,
            HelpText = "Sets process to use a console log sink for application logging.\n"
                       + "Parameter to this option must be valid LogLevel values = (Verbose,Debug,Information,Warning,Error,Fatal)\n"
                       + "Note: consoleLog, eventLog and fileLog options can be combined.\n"
                       + "Note 2: If neither consoleLog, eventLog nor fileLog options are supplied the default (in settings) logging configuration is used.\n"
                       + "Example: '--consoleLog warning'")]
        public LogEventLevel ConsoleLogConfig { get; set; } = LogEventLevel.Information;
    }
}
