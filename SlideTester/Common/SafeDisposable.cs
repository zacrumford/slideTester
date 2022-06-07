using System;
using System.Threading;

namespace SlideTester.Common;

/// <summary>
/// This class implements the IDisposable interface for releasing managed and 
/// native resources.  Derived classes SHOULD NOT implement a finalizer without
/// following the comments for ~SafeDisposable().
/// </summary>
/// <remarks>For more information, see the following MSDN articles related to IDisposable:
/// 
/// IDisposable Interface
/// http://msdn.microsoft.com/en-us/library/system.idisposable.aspx
/// 
/// CA1063: Implement IDisposable correctly
/// http://msdn.microsoft.com/en-us/library/ms244737.aspx
/// 
/// </remarks>
public abstract class SafeDisposable : IDisposable
{
    /// <summary>
    /// A null safe dispose call for any IDisposable based object
    /// This call executes the dispose method on a worker thread and waits a given
    /// time for the dispose method to complete. If a timeout occurs or no timeout was specified 
    /// then this method returns false, else it will return true.
    /// </summary>
    /// <param name="obj">object to dispose</param>
    /// <param name="timeout">maximum amount of time we are willing to wait for the object to be dispose. If null then no wait is performed</param>
    /// <returns>if we waited until the disposal thread successfully executed then true, if our wait timed out or no wait duration was supplied then false</returns>
    public static bool DisposeOffThread(IDisposable? obj, TimeSpan? timeout)
    {
        bool disposed = false;

        if (obj != null)
        {
            Thread workerThread = new Thread(obj.Dispose)
            {
                IsBackground = true // allows app to close even if thread is still executing (this is our dtor hang protection)
            };
            workerThread.Start();

            if (timeout != null)
            {
                disposed = workerThread.Join(Convert.ToInt32(timeout.Value.TotalMilliseconds)); // wait for either the thread to finish or the timeout to elapse
            }
        }

        return disposed;
    }

    /// <summary>
    /// A null safe dispose call for any IDisposable based object
    /// </summary>
    /// <param name="obj"></param>
    public static void Dispose( IDisposable obj )
    {
        if ( obj != null )
        {
            obj.Dispose( );
        }
    }

    private long m_lock = 0;
    private bool m_isDisposed = false;

    public SafeDisposable()
    {
    }

    /// <summary>
    /// Finalizer for this class.  If overriden in a derived class, the derived class
    /// MUST invoke this destructor explicitly.  Safest design is to NOT declare
    /// destructors in derived classes and defer all cleanup to the 
    /// CleanupDisposableObjects and CleanupUnmanagedResources methods below.
    /// </summary>
    ~SafeDisposable()
    {
        DisposeProxy(false);
    }

    #region IDisposable Members

    /// <summary>
    /// The Dispose method may be called explicitly or implicitly (via the using keyword)
    /// multiple times.
    /// </summary>
    public void Dispose()
    {
        // Invoke DisposProxy with releaseManagedResources = true.

        DisposeProxy(true);

        // Suppress finalization for this object.  Since unmanaged 
        // resources have been released in the above DisposeProxy 
        // call, there's no need for the finalizer to run on this
        // object.

        GC.SuppressFinalize(this);
    }

    #endregion //IDisposable

    /// <summary>
    /// DisposeProxy ensures that the right types of resources are released at 
    /// the right time.
    /// 
    /// If called by the finalizer, it is NOT safe to release managed resources,
    /// but it is safe to release unmanaged resources.
    /// 
    /// If called explicitly or implicitly via Dispose, it is safe to release
    /// both managed and unmanaged resources.
    /// 
    /// Typical sequence is one of the following scenarios.
    /// 
    /// 1. Dispose called called by client code, finalizer invoked later on the
    /// GC thread.
    /// 
    /// 2. Dispose is not called by the client code, finalizer invoked eventually
    /// on the GC thread.
    /// </summary>
    /// <param name="releaseManagedResources">True if called from the public Dispose 
    /// method.  False if called from the finalizer by the GC thread.  If false,
    /// do not release managed resources.</param>
    protected void DisposeProxy(bool releaseManagedResources)
    {
        // Acquire the disposal lock for this object.

        if (1 == Interlocked.Increment(ref m_lock))
        {
            try
            {
                // If we are already disposed, don't do anything.

                if (!m_isDisposed)
                {
                    // Have not been disposed, attempt cleanup.

                    try
                    {
                        try
                        {
                            // Attempt to clean up managed resources.

                            if (releaseManagedResources)
                            {
                                CleanupDisposableObjects();
                            }
                        }
                        finally
                        {
                            // Attempt to clean up unmanaged resources.

                            CleanupUnmanagedResources();
                        }
                    }
                    finally
                    {
                        // Mark this as disposed, regardless of whether or not disposal
                        // triggered an exception.

                        m_isDisposed = true;
                    }
                }
            }
            finally
            {
                // Release the disposal lock.

                Interlocked.Decrement(ref m_lock);
            }
        }
    }

    /// <summary>
    /// Boolean flag indicating whether or not this object has been disposed.
    /// </summary>
    /// <param name="instance">Object instance in question.</param>
    /// <returns>True if the object is null, is not alive, or has been 
    /// disposed.  False if the object is not null, is alive, and has not 
    /// been disposed.</returns>
    public static bool IsDisposed(
        SafeDisposable instance)
    {
        bool result = true;
        if (null != instance)
        {
            WeakReference weakRef = new WeakReference(instance);
            if (weakRef.IsAlive)
            {
                result = instance.m_isDisposed;
            }
        }
        return result;
    }

    /// <summary>
    /// Boolean flag indicating whether or not this object has been disposed.
    /// </summary>
    /// <param name="instance">Object instance in question.</param>
    /// <returns>True if the object is null, or is not alive.  False if 
    /// the object is not null and is alive.</returns>
    public static bool IsDisposed(
        object instance)
    {
        bool result = true;
        if (null != instance)
        {
            WeakReference weakRef = new WeakReference(instance);
            result = !weakRef.IsAlive;
        }
        return result;
    }

    /// <summary>
    /// Method to clean up managed resources.  This method will be invoked 
    /// once and only once, and only if there is an explicit dispose.
    /// </summary>
    protected abstract void CleanupDisposableObjects();

    /// <summary>
    /// Method to clean up native resources.  This method will be invoked
    /// once and only once, regardless of if there is an explicit dispose.
    /// </summary>
    protected abstract void CleanupUnmanagedResources();
}

