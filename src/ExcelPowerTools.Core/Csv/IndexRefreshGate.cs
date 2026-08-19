using System;
using System.Threading;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>
    /// Rejects nested index refreshes. Excel can synchronously raise activation events while a
    /// worksheet is being added, so event suppression alone is not a sufficient recursion guard.
    /// </summary>
    public sealed class IndexRefreshGate
    {
        private int _entered;

        public bool IsEntered => Volatile.Read(ref _entered) != 0;

        /// <summary>Returns a lease for the first caller, or null for a nested/concurrent caller.</summary>
        public IDisposable TryEnter()
        {
            return Interlocked.CompareExchange(ref _entered, 1, 0) == 0
                ? new Lease(this)
                : null;
        }

        private sealed class Lease : IDisposable
        {
            private IndexRefreshGate _owner;

            public Lease(IndexRefreshGate owner) => _owner = owner;

            public void Dispose()
            {
                IndexRefreshGate owner = Interlocked.Exchange(ref _owner, null);
                if (owner != null) Volatile.Write(ref owner._entered, 0);
            }
        }
    }
}
