using System.Threading;
using System.Threading.Tasks;

using SlideTester.App.CommandLine;

namespace SlideTester.App
{
    interface IVerbProcessor
    {
        Task DoWork(CancellationToken token);        
    }
}