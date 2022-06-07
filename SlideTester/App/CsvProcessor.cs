using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;

using SlideTester.App.CommandLine;
using SlideTester.Common;

namespace SlideTester.App
{
    internal class CsvProcessor : BaseProcessor<CsvOptions>
    {
        // ReSharper disable once ClassNeverInstantiated.Local
        private class CsvProcessorRecord
        {
            [Index(0)]
            public string SlideDeckUri { get; set; } = string.Empty;
            
            [Index(1)]
            public string AwsRegion { get; set; } = string.Empty;
        }
        
        public CsvProcessor(
            CsvOptions options) : base(options)
        {
            
        }

        public override async Task DoWork(CancellationToken token)
        {
            Console.WriteLine($"Starting csv ({this.Options.CsvFilePath}) processing");
            
            ChkArg.FileExists(this.Options.CsvFilePath, nameof(this.Options.CsvFilePath));
            ChkArg.IsNotNull(token,nameof(token));
            token.ThrowIfCancellationRequested();
            List<CsvProcessorRecord> csvProcessorRecords;
            
            Console.WriteLine($"Parsing csv file.");
            
            CsvConfiguration configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = this.Options.HasHeaderRecord,
            };
            
            using (StreamReader streamReader = new StreamReader(this.Options.CsvFilePath))
            using (CsvReader csv = new CsvReader(streamReader, configuration))
            {
                await csv.ReadAsync().ConfigureAwait(false);
                csvProcessorRecords = csv.GetRecords<CsvProcessorRecord>().ToList();  
            }
            
            Console.WriteLine($"Csv parsing complete. Records found: {csvProcessorRecords.Count}");

            Console.WriteLine($"Processing records one by one.");
            
            for (int i = 0; i < csvProcessorRecords.Count; ++i)
            {
                Console.WriteLine($"Processing slide deck {i+1}/{csvProcessorRecords.Count}");
                await this.ProcessFile(
                    csvProcessorRecords[i].SlideDeckUri,
                    csvProcessorRecords[i].AwsRegion,
                    token).ConfigureAwait(false);
            }
            
            Console.WriteLine($"Processing of all csv records is complete.");
        }
    }
}