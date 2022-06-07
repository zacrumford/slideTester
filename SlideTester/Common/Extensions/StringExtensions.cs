using System.IO;
using System.Text;

namespace SlideTester.Common.Extensions;
    
public static class StringExtensions
{
    public static Stream ToStream(this string instance, Encoding encoding)
    {
        return new MemoryStream(encoding.GetBytes(instance));
    }
}
