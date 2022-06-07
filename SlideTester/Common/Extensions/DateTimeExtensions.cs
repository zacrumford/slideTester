using System;
using System.Globalization;

namespace SlideTester.Common.Extensions;

public static class DateTimeExtensions
{
    public static DateTime UnixEpoch => new DateTime(1970, 1, 1, 0, 0, 0);
    
    public static long ToUnixTime(
        this DateTime dateTime)
    {
        var timeSpan = (dateTime - UnixEpoch);
        return (long)timeSpan.TotalSeconds;
    }

    public static DateTime FromUnixTime(
        long unixTime)
    {
        return UnixEpoch + TimeSpan.FromSeconds(unixTime);
    }

    public static DateTime ParseToUtc(string date, string format)
    {
        return DateTime.SpecifyKind(
            DateTime.ParseExact(date,format, CultureInfo.InvariantCulture), 
            DateTimeKind.Utc);
    }
}
