using System;
using System.Threading;
using System.Threading.Tasks;

using SlideTester.App.CommandLine;

namespace SlideTester.App
{
    internal class UriProcessor : BaseProcessor<UriOptions>
    {
        public UriProcessor(
            UriOptions options) : base(options)
        {
            
        }

        public override async Task DoWork(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            Console.WriteLine($"Processing of single PPT file starting.");
            
            await this.ProcessFile(this.Options.PowerpointFileUri, this.Options.Region, token)
                .ConfigureAwait(false);

            Console.WriteLine($"All work complete.");
        }
    }
}