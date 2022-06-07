using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

// ReSharper disable ConditionIsAlwaysTrueOrFalse

namespace SlideTester.Common;

/// <summary>
/// Helper class for checking args and throwing errors if they are not expected.
/// We should add to this class as more of these arg check types are written throughout the code.
/// </summary>
public static class ChkArg
{
    public static void IsTrue(bool arg, string failureText, params object[] fmtArgs)
    {
        if (!arg)
        {
            throw new ArgumentException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsFalse(bool arg, string failureText, params object[] fmtArgs)
    {
        if (arg)
        {
            throw new ArgumentException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotNull(object obj, string failureText, params object[] fmtArgs)
    {
        if (null == obj)
        {
            throw new ArgumentNullException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotNull(object obj)
    {
        if (null == obj)
        {
            throw new ArgumentNullException();
        }
    }

    public static void IsNotGuidEmpty(Guid guid, string failureText, params object[] fmtArgs)
    {
        if (Guid.Empty == guid)
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotGuidEmpty(Guid guid)
    {
        if (Guid.Empty == guid)
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    /// <summary>
    /// Throws an exception if the string is null or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    /// <param name="failureText">failure message and format string</param>
    /// <param name="fmtArgs">format string args</param>
    public static void IsNotNullOrEmpty(string arg, string failureText, params object[] fmtArgs)
    {
        if (string.IsNullOrEmpty(arg))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the string is null or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    public static void IsNotNullOrEmpty(string arg)
    {
        if (string.IsNullOrEmpty(arg))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    /// <summary>
    /// Throws an exception if the string is null, whitespace or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    public static void IsNotNullOrWhiteSpace(string arg)
    {
        if (null == arg)
        {
            throw new ArgumentNullException();
        }
        else if (string.IsNullOrWhiteSpace(arg)) // double checks for null but who cares ;)
        {
            throw new ArgumentOutOfRangeException();
        }
    }
    /// <summary>
    /// Throws an exception if the string is null, whitespace or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    /// <param name="failureText">failure message and format string</param>
    /// <param name="fmtArgs">format string args</param>
    public static void IsNotNullOrWhiteSpace(string arg, string failureText, params object[] fmtArgs)
    {
        if (null == arg)
        {
            throw new ArgumentNullException(string.Format(failureText, fmtArgs));
        }
        else if (string.IsNullOrWhiteSpace(arg)) // double checks for null but who cares ;)
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the array is null or empty
    /// </summary>
    /// <param name="arg">array to test</param>
    /// <param name="failureText">failure message and format string</param>
    /// <param name="fmtArgs">format string args</param>
    public static void IsNotNullOrEmpty<T>(IEnumerable<T> arg, string failureText, params object[] fmtArgs)
    {
        if (null == arg)
        {
            throw new ArgumentNullException(string.Format(failureText, fmtArgs));
        }
        else if (!arg.Any())
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the array is null or empty
    /// </summary>
    /// <param name="arg">array to test</param>
    public static void IsNotNullOrEmpty<T>(IEnumerable<T> arg)
    {
        if (null == arg)
        {
            throw new ArgumentNullException();
        }
        else if (!arg.Any())
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    /// <summary>
    /// Throws an exception if the array contains items
    /// </summary>
    /// <param name="arg">array to test</param>
    /// <param name="failureText">failure message and format string</param>
    /// <param name="fmtArgs">format string args</param>
    public static void IsNullOrEmpty<T>(IEnumerable<T> arg, string failureText, params object[] fmtArgs)
    {
        if (null != arg)
        {
            throw new ArgumentNullException(string.Format(failureText, fmtArgs));
        }
        
        if (arg != null && arg.Any())
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the array contains items
    /// </summary>
    /// <param name="arg">array to test</param>
    public static void IsNullOrEmpty<T>(IEnumerable<T> arg)
    {
        if (null != arg)
        {
            throw new ArgumentNullException();
        }
        
        if (arg != null && arg.Any())
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    /// <summary>
    /// Checks that the argument is the expected type and is non-null.
    /// </summary>
    /// <typeparam name="TType">expected type</typeparam>
    /// <param name="arg">object value</param>
    /// <returns>the cast object</returns>
    public static TType CastIsNotNull<TType>(object arg)
        where TType : class
    {
        if (arg == null)
        {
            throw new ArgumentNullException(nameof(arg));
        }
        
        TType result = arg as TType ?? throw new InvalidOperationException(
            String.Format("Expected non-null instance of {0}", typeof(TType).Name));
        
        return result;
    }

    /// <summary>
    /// Checks that the argument is either null or the expected type.
    /// </summary>
    /// <typeparam name="TType">expected type</typeparam>
    /// <param name="arg">object value</param>
    /// <returns>the cast object</returns>
    public static TType CastOrNull<TType>(object arg)
        where TType : class
    {
        if (arg == null)
        {
            throw new ArgumentNullException(nameof(arg));
        }
        
        TType result = arg as TType ?? throw new InvalidOperationException(
            String.Format("Expected null or instance of {0}", typeof(TType).Name));

        return result;
    }

    public static void AreEqual<T>(T obj, T objExpectedValue, string failureText, params object[] fmtArgs)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (!obj.Equals(objExpectedValue))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void AreNotEqual<T>(T obj, T objExpectedValue, string failureText, params object[] fmtArgs)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (obj.Equals(objExpectedValue))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void AreEqual<T>(T obj, T objExpectedValue)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (!obj.Equals(objExpectedValue))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void AreNotEqual<T>(T obj, T objExpectedValue)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (obj.Equals(objExpectedValue))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void IsLessThan<T>(T lhs, T rhs, string failureText, params object[] fmtArgs) where T : IComparable<T>
    {
        if (0 <= lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsLessThan<T>(T lhs, T rhs) where T : IComparable<T>
    {
        if (0 <= lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void IsLessThanOrEqualTo<T>(T lhs, T rhs, string failureText, params object[] fmtArgs) where T : IComparable<T>
    {
        if (0 < lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsLessThanOrEqualTo<T>(T lhs, T rhs) where T : IComparable<T>
    {
        if (0 < lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void IsGreaterThan<T>(T lhs, T rhs, string failureText, params object[] fmtArgs) where T : IComparable<T>
    {
        if (0 >= lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsGreaterThan<T>(T lhs, T rhs) where T : IComparable<T>
    {
        if (0 >= lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void IsGreaterThanOrEqualTo<T>(T lhs, T rhs, string failureText, params object[] fmtArgs) where T : IComparable<T>
    {
        if (0 > lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsGreaterThanOrEqualTo<T>(T lhs, T rhs) where T : IComparable<T>
    {
        if (0 > lhs.CompareTo(rhs))
        {
            throw new ArgumentOutOfRangeException();
        }
    }

    public static void IsInRangeInclusive<T>(T lhs, T lbound, T ubound, string failureText, params object[] fmtArgs) where T : IComparable<T>
    {
        if ((lhs.CompareTo(lbound) < 0) || (lhs.CompareTo(ubound) > 0))
        {
            throw new ArgumentOutOfRangeException(string.Format(failureText, fmtArgs));
        }
    }

    public static bool SucceededHResult(int hr)
    {
        return hr >= 0;
    }

    public static bool FailedHResult(int hr)
    {
        return hr < 0;
    }

    // BUGBUG: 35627: SafeBatchDeleter::SafeDeleteInBatches does not appear to do arg checking correctly
    public static void StringMatchesFormatArgs(string str, object[] formatArgs, string failureText)
    {
        try
        {
            string pattern = "{([0-9]+)}";
            MatchCollection mc = Regex.Matches(str, pattern);
            int place;
            HashSet<int> seenInts = new HashSet<int>();
            foreach (Match match in mc)
            {
                if (!int.TryParse(match.Groups[1].Value, out place))
                {
                    throw new Exception();
                }
                else
                {
                    seenInts.Add(place);
                }
            }

            int[] seenIntsArray = seenInts.ToArray();
            Array.Sort(seenIntsArray);

            AreEqual(seenIntsArray.Count() - 1, seenIntsArray[seenIntsArray.Count() - 1]);
        }
        catch (Exception e)
        {
            throw new ArgumentException(failureText, e);
        }
    }

/// <summary>
        /// Checks to see if a file exists, if it does not then this method throws
        /// </summary>
        public static void FileExists(
            string filename,
            string failureText,
            params object[] formatArgs)
        {
            ChkExpect.IsNotNullOrEmpty(filename, nameof(filename));

            if (!System.IO.File.Exists(filename))
            {
                if (failureText == null)
                {
                    throw new ArgumentException($"The file does not exist: {filename}");
                }
                else
                {
                    throw new ArgumentException(
                        $"File was expected to exist yet none was found at location ({filename}).\n" +
                        string.Format(failureText, formatArgs));
                }
            }
        }

        /// <summary>
        /// Checks to see if a file does not exist, if it does then this method throws
        /// </summary>
        public static void FileDoesNotExist(
            string filename,
            string failureText,
            params object[] formatArgs)
        {
            ChkExpect.IsNotNullOrEmpty(filename, nameof(filename));

            if (System.IO.File.Exists(filename))
            {
                if (string.IsNullOrEmpty(failureText))
                {
                    throw new ArgumentException($"The file exists but was expected not to exist: {filename}");
                }
                else
                {
                    throw new ArgumentException(
                        $"File was expected to not exist yet but one already exists at location ({filename}).\n" +
                        string.Format(failureText, formatArgs));
                }
            }
        }
        
        /// <summary>
        /// Checks to see if a directory exists, if it does not then this method throws
        /// </summary>
        public static void DirectoryExists(
            string directoryPath, 
            string failureText, 
            params object[] formatArgs)
        {
            ChkExpect.IsNotNullOrEmpty(directoryPath, nameof(directoryPath));

            if (!System.IO.Directory.Exists(directoryPath))
            {
                if (string.IsNullOrEmpty(failureText))
                {
                    throw new ArgumentException($"The directory does not exist: {directoryPath}");
                }
                else
                {
                    throw new ArgumentException(
                        $"The directory was expected to exist yet none was found at location ({directoryPath}).\n" +
                        string.Format(failureText, formatArgs));
                }
            }
        }

        /// <summary>
        /// Checks to see if a directory does not exist, if it does then this method throws
        /// </summary>
        public static void DirectoryDoesNotExist(
            string directoryPath, 
            string failureText, 
            params object[] formatArgs)
        {
            ChkExpect.IsNotNullOrEmpty(directoryPath, nameof(directoryPath));

            if (System.IO.Directory.Exists(directoryPath))
            {
                if (string.IsNullOrEmpty(failureText))
                {
                    throw new ArgumentException($"The directory does not exist: {directoryPath}");
                }
                else
                {
                    throw new ArgumentException(
                        $"The directory was expected to not exist yet but one already exists at location ({directoryPath}).\n" +
                        string.Format(failureText, formatArgs));                    
                }
            }
        }
}

public static class ChkExpect
{
    /// <summary>
    /// Method to help quickly throw an error when we the caller knows we are in a failure scenario
    /// </summary>
    /// <param name="failureText">failure string to insert in exception</param>
    /// <param name="fmtArgs">optional failure message args</param>
    public static void Fail(string failureText, params object[] fmtArgs)
    {
        throw new InvalidOperationException(string.Format(failureText, fmtArgs));
    }

    public static void AreEqual<T>(T obj, T objExpectedValue, string failureText, params object[] fmtArgs)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (!obj.Equals(objExpectedValue))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void AreEqual<T>(T obj, T objExpectedValue)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (!obj.Equals(objExpectedValue))
        {
            throw new InvalidOperationException(string.Format("AreEqual check failed for values: \"{0}\" & \"{1}\"", obj, objExpectedValue));
        }
    }

    public static void AreNotEqual<T>(T obj, T objExpectedValue, string failureText, params object[] fmtArgs)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (obj.Equals(objExpectedValue))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void AreNotEqual<T>(T obj, T objExpectedValue)
    {
        if (obj == null)
        {
            throw new ArgumentNullException(nameof(obj));
        }
        
        if (obj.Equals(objExpectedValue))
        {
            throw new InvalidOperationException(string.Format("AreNotEqual check failed for values: \"{0}\" & \"{1}\"", obj, objExpectedValue));
        }
    }

    public static void IsTrue(bool result, string failureText, params object[] fmtArgs)
    {
        if (!result)
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsTrue(bool result)
    {
        if (!result)
        {
            throw new InvalidOperationException("Expression is not true, this result is unexpected.");
        }
    }

    public static void IsNull<T>(T value)
    {
        if (null != value)
        {
            throw new InvalidOperationException("value is expected to be null");
        }
    }

    public static void IsNull<T>(T value, string failureText, params object[] fmtArgs)
    {
        if (null != value)
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotNull<T>(T value)
    {
        if (null == value)
        {
            throw new InvalidOperationException("value is expected to be non-null");
        }
    }

    public static void IsNotNull<T>(T value, string failureText, params object[] fmtArgs)
    {
        if (null == value)
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotGuidEmpty(Guid guid, string failureText, params object[] fmtArgs)
    {
        if (Guid.Empty == guid)
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsNotNullOrGuidEmpty(Guid? guid, string failureText, params object[] fmtArgs)
    {
        IsNotNull(guid, failureText, fmtArgs);
        if (guid.HasValue)
        {
            IsNotGuidEmpty(guid.Value, failureText, fmtArgs);
        }
    }

    /// <summary>
    /// Throws an exception if the string is null or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    /// <param name="failureText">failure message</param>
    /// <param name="fmtArgs">Format args for failureText</param>
    public static void IsNotNullOrEmpty(string arg, string failureText, params object[] fmtArgs)
    {
        if (string.IsNullOrEmpty(arg))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the string is null or empty
    /// </summary>
    /// <param name="arg">string to test</param>
    /// <param name="failureText">optional failure message</param>
    /// <param name="fmtArgs">Format args for failureText</param>
    public static void IsNotNullOrWhiteSpace(string arg, string failureText, params object[] fmtArgs)
    {
        if (string.IsNullOrWhiteSpace(arg))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the string is not null and not empty
    /// </summary>
    /// <param name="arg">string to test</param>
    /// <param name="failureText">optional failure message</param>
    /// <param name="fmtArgs">Format args for failureText</param>
    public static void IsNullOrEmpty(string arg, string failureText, params object[] fmtArgs)
    {
        if (!string.IsNullOrEmpty(arg))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the array is null or empty
    /// </summary>
    /// <param name="arg">array to test</param>
    /// <param name="failureText">optional failure message</param>
    /// <param name="fmtArgs">Format args for failureText</param>
    public static void IsNotNullOrEmpty<T>(IEnumerable<T> arg, string failureText, params object[] fmtArgs)
    {
        if (null == arg || !arg.Any())
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Throws an exception if the array contains items
    /// </summary>
    /// <param name="arg">array to test</param>
    /// <param name="failureText">optional failure message</param>
    /// <param name="fmtArgs">Format args for failureText</param>
    public static void IsNullOrEmpty(Array arg, string failureText, params object[] fmtArgs)
    {
        if (!IsNullOrEmpty(arg))
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    /// <summary>
    /// Determines whether the array is either null or empty
    /// </summary>
    /// <param name="array">the array</param>
    /// <returns>true if the array is null or empty</returns>
    public static bool IsNullOrEmpty(Array array)
    {
        return (array == null || array.Length == 0);
    }

    public static void IsFalse(bool result, string failureText, params object[] fmtArgs)
    {
        if (result)
        {
            throw new InvalidOperationException(string.Format(failureText, fmtArgs));
        }
    }

    public static void IsFalse(bool result)
    {
        if (result)
        {
            throw new InvalidOperationException("Expression is true, this result is unexpected.");
        }
    }
}
