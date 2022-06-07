using System;

namespace SlideTester.Common.Extensions;

public static class IntExtensions
{
    /// <summary>
    /// Simple helper method to get the greatest common divisor between two numbers
    /// This is helpful to compute the aspect ratio.
    /// If numbers are co-prime then GCD is 1
    /// </summary>
    /// <param name="first"></param>
    /// <param name="second"></param>
    /// <returns>The GCD between the two parameters</returns>
    public static int GreatestCommonDivisor(
        this int first,
        int second)
    {
        int gcd = 0;

        for (int i = Math.Min(first, second); i > 0; --i)
        {
            if ((first % i) == 0 && (second % i) == 0)
            {
                gcd = i;
                break;
            }
        }

        return gcd;
    }
}
