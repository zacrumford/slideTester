using System;
using System.Collections.Generic;
using System.Linq;

namespace SlideTester.Common.Extensions;

public static class IEnumerableExtensions
{
    /// <summary>
    /// Return the median of a sequence; returns 0 for an empty sequence
    /// </summary>
    /// <param name="sequence"></param>
    /// <returns>The median value, averaged between two nearest to center for an even length sequence</returns>
    public static double Median(this IEnumerable<double> sequence)
    {
        if (!sequence.Any())
        {
            return 0;
        }
        int numberOfGaps = sequence.Count() - 1;
        return sequence.OrderBy(s => s).Skip(numberOfGaps / 2).Take(1 + (numberOfGaps % 2)).Average();
    }

    /// <summary>
    /// Creates an infinite output enumerable for the input enumerable by
    /// re-looping over the enumerable when it has reached the end.
    /// <remarks>
    /// Performs multiple enumerations of the input enumerable. Unsafe for yielding enumerators.
    /// </remarks> 
    /// </summary>
    /// <returns>An enumerable which will yield infinite results from the input enumerable</returns>
    public static IEnumerator<T> Infinite<T>(
        this IEnumerable<T> enumerable)
    {
        while (true)
        {
            foreach (var pri in enumerable)
            {
                yield return pri;
            }
        }
    }

    /// <summary>
    /// True or false for an IEnumerable of T is sorted, where T : IComparable
    /// </summary>
    /// <typeparam name="T">Any comparable object</typeparam>
    /// <param name="sequence">A sequence of comparable objects of type T</param>
    /// <param name="descending">When true, asserts descending order. Else asserts ascending order</param>
    /// <returns>true if sorted, else false</returns>
    public static bool IsSorted<T>(this IEnumerable<T> sequence, bool descending=false) where T : IComparable
    {
        if (!sequence.Any())
        {
            return true;
        }

        int comparabilityFactor = descending ? -1 : 1;
        T previousEntry = sequence.First();
        foreach (T entry in sequence.Skip(1))
        {
            if (comparabilityFactor * previousEntry.CompareTo(entry) > 0)
            {
                return false;
            }
            previousEntry = entry;
        }
        return true;
    }

    /// <summary>
    /// Concatenates one or more enumerables on to a target enumerable
    /// </summary>
    public static IEnumerable<T> Concatenate<T>(
        this IEnumerable<T> val,
        params IEnumerable<T>[] listOfLists)
    {
        return val.Concat(listOfLists.SelectMany(str => str));
    }

    /// <summary>
    /// Combines an enumerable of enumerables into a single enumerable
    /// </summary>
    public static IEnumerable<T> Combine<T>(
        this IEnumerable<IEnumerable<T>> val)
    {
        return val.SelectMany(str => str);
    }

    /// <summary>
    /// Does enumerable a contain all of enumerable b
    /// </summary>
    public static bool ContainsAllItems<T>(this IEnumerable<T> a, IEnumerable<T> b)
    {
        return !b.Except(a).Any();
    }
    
    /// <summary>
    /// Does enumerable a contain none of enumerable b
    /// </summary>
    public static bool ContainsNone<T>(this IEnumerable<T> a, IEnumerable<T> b)
    {
        return !a.Intersect(b).Any();
    }
}
