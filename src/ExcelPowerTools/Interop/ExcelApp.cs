using System;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>Convenience access to the running Excel Application as a late-bound COM object.</summary>
    internal static class ExcelApp
    {
        /// <summary>The Excel Application, late-bound (use like VBA: app.Selection, app.ActiveSheet, ...).</summary>
        public static dynamic App => ExcelDnaUtil.Application;
    }

    /// <summary>
    /// Saves and restores Application.ScreenUpdating / EnableEvents / Calculation around a batch of work,
    /// the way a well-behaved VBA macro does. Every access is defensive because these properties throw
    /// when no workbook is open.
    /// </summary>
    internal sealed class ExcelStateGuard : IDisposable
    {
        private readonly dynamic _app;
        private readonly bool _screenUpdating;
        private readonly bool _enableEvents;
        private readonly object _calculation;
        private readonly bool _calcCaptured;

        public ExcelStateGuard(bool manualCalculation = false, bool suppressEvents = false)
        {
            _app = ExcelDnaUtil.Application;

            _screenUpdating = TryGetBool(() => _app.ScreenUpdating, true);
            _enableEvents = TryGetBool(() => _app.EnableEvents, true);

            _calcCaptured = TryGet(() => (object)_app.Calculation, out _calculation);

            TrySet(() => _app.ScreenUpdating = false);
            if (suppressEvents) TrySet(() => _app.EnableEvents = false);
            if (manualCalculation && _calcCaptured) TrySet(() => _app.Calculation = Xl.xlCalculationManual);
        }

        public void Dispose()
        {
            TrySet(() => _app.ScreenUpdating = _screenUpdating);
            TrySet(() => _app.EnableEvents = _enableEvents);
            if (_calcCaptured) TrySet(() => _app.Calculation = _calculation);
        }

        private static bool TryGetBool(Func<bool> getter, bool fallback)
        {
            try { return getter(); } catch { return fallback; }
        }

        private static bool TryGet<T>(Func<T> getter, out T value)
        {
            try { value = getter(); return true; }
            catch { value = default; return false; }
        }

        private static void TrySet(Action setter)
        {
            try { setter(); } catch { /* no workbook / not applicable */ }
        }
    }
}
