using System;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class IndexRefreshGateTests
    {
        [Fact]
        public void NestedEntryIsRejected()
        {
            var gate = new IndexRefreshGate();

            using (IDisposable first = gate.TryEnter())
            {
                Assert.NotNull(first);
                Assert.True(gate.IsEntered);
                Assert.Null(gate.TryEnter());
            }
        }

        [Fact]
        public void DisposeAllowsNextIndependentEntry()
        {
            var gate = new IndexRefreshGate();

            gate.TryEnter().Dispose();

            using (IDisposable next = gate.TryEnter()) Assert.NotNull(next);
        }

        [Fact]
        public void ExceptionUnwindingReleasesGate()
        {
            var gate = new IndexRefreshGate();

            Assert.Throws<InvalidOperationException>(() => ThrowInsideGate(gate));

            Assert.False(gate.IsEntered);
            using (IDisposable next = gate.TryEnter()) Assert.NotNull(next);
        }

        [Fact]
        public void SimulatedActivationDuringRefreshDoesNotReenter()
        {
            var gate = new IndexRefreshGate();
            int refreshCount = 0;

            Action? refresh = null;
            refresh = () =>
            {
                IDisposable lease = gate.TryEnter();
                if (lease == null) return;
                using (lease)
                {
                    refreshCount++;
                    refresh!(); // simulates a synchronous SheetActivate callback
                }
            };

            refresh();

            Assert.Equal(1, refreshCount);
        }

        private static void ThrowInsideGate(IndexRefreshGate gate)
        {
            using (IDisposable lease = gate.TryEnter())
            {
                Assert.NotNull(lease);
                throw new InvalidOperationException("test");
            }
        }
    }
}
