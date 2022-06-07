using System;
using System.Diagnostics;
using System.IO;

using SlideTester.Common.Log;

namespace SlideTester.Common;

/// <summary>
/// class for getting unique scratch folders, then cleaning up after them once you're done.
/// </summary>
public class ScratchSpace : SafeDisposable
{
    public static string ScratchBasePath { get; set; } = "c:\\panopto\\scratch\\"; 
    public static string GetNewPath()
    {
        return GetNewPath(
            baseFolder: ScratchBasePath,
            pid: Process.GetCurrentProcess().Id);
    }
    
    public static string GetNewPath(string baseFolder, int pid)
    {
        long ticks = DateTime.UtcNow.Ticks;
        
        string output = System.IO.Path.Combine(
            baseFolder,
            $"{Guid.NewGuid().ToString()}_{Environment.CurrentManagedThreadId:x}_{pid}");

        return output;
    }

    /// <summary>
    /// Extracts the ID of the process that created the scratch folder. 
    /// </summary>
    /// <param name="path">The path to the scratch folder.</param>
    /// <returns>The ID of the process that created the scratch folder, or null if the path did not conform
    /// to expectations.</returns>
    public static int? GetPidFromScratchPath(string path)
    {
        ChkArg.IsNotNull(path, nameof(path));

        int? output = null;
        int idx = path.LastIndexOf("_", StringComparison.InvariantCultureIgnoreCase);
        
        if (idx > 0
            && idx < path.Length - 1
            && int.TryParse(path.Substring(idx + 1), out int pid))
        {
            output = pid;
        }

        return output;
    }

    /// <summary>
    /// The path that was chosen.
    /// </summary>
    public string Path { get; private set; }

    public ScratchSpace()
    {
        this.Path = GetNewPath();

        // Create the folder.
        try
        {
            Directory.CreateDirectory(this.Path);
        }
        catch (Exception e)
        {
            // Rethrow with failure message plus detail information (inner exception).
            throw new IOException("Failed to create scratch space: " + this.Path, e);
        }
    }

    /// <summary>
    /// get the full path to a location rooted underneath this scratch folder.
    /// </summary>
    public string GetFullPath(string relativePath)
    {
        return System.IO.Path.Combine(this.Path, relativePath);
    }

    /// <summary>
    /// Creates a new unique folder under the scratch space root. Returns the full path to the caller.
    /// </summary>
    /// <returns>Full path of newly created folder</returns>
    public string CreateUniqueFolder()
    {
        string newDir = System.IO.Path.Combine(this.Path, Guid.NewGuid().ToString()).ToLowerInvariant();

        // Create the folder.
        try
        {
            Directory.CreateDirectory(newDir);
        }
        catch (Exception ex)
        {
            // Rethrow with failure message plus detail information (inner exception).
            throw new IOException($"Failed to new directory under scratch space: {newDir}", ex);
        }

        return newDir;
    }

    /// <summary>
    /// Generates a unique filename pathed under this scratch space and 
    /// returns the full path of the file. Note: no actual file is written to
    /// disk. This method is designed to be used when we need to generate a
    /// filename which may get created by other processing but the file should be cleaned
    /// up when the scrach object is being cleaned up.
    /// </summary>
    /// <returns>Generated unique filename w/ full path. Filename will use a .tmp extension</returns>
    public string UniqueFilePath()
    {
        return this.UniqueFilePath("tmp");
    }

    /// <summary>
    /// Generates a unique filename pathed under this scratch space and 
    /// returns the full path of the file. Note: no actual file is written to
    /// disk. This method is designed to be used when we need to generate a
    /// filename which may get created by other processing but the file should be cleaned
    /// up when the scrach object is being cleaned up.
    /// </summary>
    /// <param name="fileExtension">Extention of the generated filename</param>
    /// <returns>Generated unique filename w/ full path</returns>
    public string UniqueFilePath(string fileExtension)
    {
        if (fileExtension.StartsWith("."))
        {
            return System.IO.Path.Combine(
                this.Path,
                $"{Guid.NewGuid()}{fileExtension}".ToLowerInvariant());
        }
        else
        {
            return System.IO.Path.Combine(
                this.Path,
                $"{Guid.NewGuid()}.{fileExtension}".ToLowerInvariant());
        }
    }

    protected override void CleanupDisposableObjects()
    {
    }

    protected override void CleanupUnmanagedResources()
    {
        try
        {
            // we ALWAYS delete even if there is not an explicit dispose.
            // (this is called in the finalizer)
            System.IO.Directory.Delete(this.Path, recursive: true);
        }
        catch (Exception ex)
        {
            // not much we can do if there are files held open by someone else.  complain.
            Logger.Write(Logs.ScratchSpaceDeleteScratchFolderFailed, this.Path, ex);
        }
    }
}

