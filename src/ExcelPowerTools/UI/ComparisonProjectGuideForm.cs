using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelPowerTools.UI
{
    internal static class ComparisonProjectSetupHelpText
    {
        internal const string OpenGuide =
            "Open the comparison-project setup guide.\r\n" +
            "You can also press F1 anywhere in this window.";
        internal const string RevisionList =
            "Select one revision to edit its label, parent, or role.\r\n" +
            "Right-click the selection for the same lineage actions and parent shortcuts.";
        internal const string MoveEarlier =
            "Move the selected revision one position earlier.\r\n" +
            "Parent links that would point forward are repaired, and confirmation is cleared.";
        internal const string MoveLater =
            "Move the selected revision one position later.\r\n" +
            "Parent links that would point forward are repaired, and confirmation is cleared.";
        internal const string SetMaster =
            "Mark the selected revision as the project master (the intended final state).\r\n" +
            "This changes its role only; it does not reorder the revision.";
        internal const string SetBaseline =
            "Mark the selected revision as an optional trusted reviewed baseline.\r\n" +
            "The baseline is reference evidence; it does not become the master automatically.";
        internal const string ClearBaseline =
            "Remove the trusted-baseline designation.\r\n" +
            "The revision remains in the project and its order and parent stay unchanged.";
        internal const string PreviousParent =
            "Set the selected revision's direct parent to the revision immediately before it.\r\n" +
            "Parentage records derivation and is separate from display order.";
        internal const string ClearParent =
            "Make the selected revision a root by removing its parent.\r\n" +
            "The current comparison requires exactly one root when you build.";
        internal const string RevisionLabel =
            "Edit the name shown for this revision in the comparison review.\r\n" +
            "This does not rename or modify the source workbook.";
        internal const string ParentRevision =
            "Choose the direct earlier revision from which the selected revision was derived.\r\n" +
            "A parent must appear earlier in the displayed order; exactly one revision must remain root.";
        internal const string InferenceEvidence =
            "Shows why Power Tools inferred this revision's order and parent.\r\n" +
            "Weak or absent evidence should be checked against your knowledge of the files.";
        internal const string MasterStatus =
            "Shows the revision currently designated as the project master (intended final state).";
        internal const string BaselineStatus =
            "Shows the optional revision currently designated as the trusted reviewed baseline.";
        internal const string OutputPath =
            "Full path for the new Review.xlsx.\r\n" +
            "Power Tools will not overwrite an existing file or use a source workbook as the destination.";
        internal const string BrowseOutput =
            "Choose a new filename and folder for Review.xlsx.\r\n" +
            "The destination must not already exist.";
        internal const string ConfirmLineage =
            "Confirm that you reviewed the order, parent links, master, and optional baseline.\r\n" +
            "Any later setup change clears this check and disables Build comparison again.";
        internal const string Limits =
            "Current alpha limits and the comparison-content rule used to bound each sheet.\r\n" +
            "Source workbooks remain read-only throughout comparison.";
        internal const string BuildComparison =
            "Compare stored values, saved calculated values, and formulas, then create Review.xlsx.\r\n" +
            "Rows and columns are aligned first; styles and other package details are not comparison changes.";
        internal const string Cancel =
            "Close setup without starting the comparison. No source workbook is changed.";
    }

    /// <summary>In-app guide for the concepts and actions in revision setup.</summary>
    internal sealed class ComparisonProjectGuideForm : Form
    {
        private static readonly Color Background = Color.FromArgb(247, 248, 252);
        private static readonly Color Surface = Color.White;
        private static readonly Color Ink = Color.FromArgb(28, 31, 42);
        private static readonly Color Muted = Color.FromArgb(96, 102, 119);
        private static readonly Color Accent = Color.FromArgb(100, 73, 214);

        private readonly RichTextBox _content = new RichTextBox();

        internal ComparisonProjectGuideForm()
        {
            Text = "Comparison Project — setup guide";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(660, 520);
            Size = new Size(820, 720);
            ShowInTaskbar = false;
            BackColor = Background;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;

            var title = new Label
            {
                Text = "Revision setup guide",
                Font = new Font("Segoe UI Semibold", 18F),
                ForeColor = Ink,
                Location = new Point(22, 18),
                AutoSize = true,
            };
            var subtitle = new Label
            {
                Text = "Turn inferred file order into an explicitly confirmed comparison lineage.",
                ForeColor = Muted,
                Location = new Point(25, 55),
                AutoSize = true,
            };

            _content.Location = new Point(22, 84);
            _content.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _content.Size = new Size(ClientSize.Width - 44, ClientSize.Height - 142);
            _content.ReadOnly = true;
            _content.BackColor = Surface;
            _content.ForeColor = Ink;
            _content.BorderStyle = BorderStyle.FixedSingle;
            _content.DetectUrls = false;
            _content.WordWrap = true;
            _content.ScrollBars = RichTextBoxScrollBars.Vertical;
            _content.AccessibleName = "Comparison project setup guide contents";

            var close = new Button
            {
                Text = "Close",
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Size = new Size(110, 34),
                Location = new Point(ClientSize.Width - 132, ClientSize.Height - 48),
                DialogResult = DialogResult.OK,
            };

            Controls.AddRange(new Control[] { title, subtitle, _content, close });
            AcceptButton = close;
            CancelButton = close;
            BuildGuide();
            Shown += (_, __) =>
            {
                _content.SelectionStart = 0;
                _content.SelectionLength = 0;
                _content.ScrollToCaret();
                close.Focus();
            };
        }

        private void BuildGuide()
        {
            AppendParagraph(
                "Use this window to correct filename and timestamp guesses before Power Tools compares the files. " +
                "The setup describes chronology and derivation; it never changes a source workbook.");

            AppendHeading("Recommended workflow");
            AppendParagraph(
                "1. Read the list from top to bottom as earlier to later. Select a row and use Move earlier or " +
                "Move later when the inferred order is wrong.\r\n" +
                "2. Review each Parent entry. The parent is the direct earlier revision from which that file was " +
                "derived; it is not merely the row above.\r\n" +
                "3. Replace unclear inferred labels with concise names that will make sense in Review.xlsx.\r\n" +
                "4. Set the master to the intended final or latest reviewed revision.\r\n" +
                "5. Optionally set a trusted baseline when one revision is known to have been reviewed carefully.\r\n" +
                "6. Choose a new Review.xlsx path, check the confirmation box, and build the comparison.");

            AppendHeading("Order, parent, root, master, and baseline");
            AppendDefinition(
                "Order",
                "The chronology shown from earlier to later. Reordering does not by itself prove that one file " +
                "was derived from another.");
            AppendDefinition(
                "Parent",
                "The direct source revision for a later revision. Parent links can describe branches when emailed " +
                "copies diverge. A parent must appear earlier than its child.");
            AppendDefinition(
                "Root",
                "A revision with no parent. The current comparison requires exactly one root for the selected set.");
            AppendDefinition(
                "Master",
                "The intended final state and primary latest revision for the project. Setting the master does not " +
                "move it or modify it.");
            AppendDefinition(
                "Trusted baseline",
                "An optional reviewed reference. It can be useful evidence, but it is not automatically the root, " +
                "parent, or master.");
            AppendDefinition(
                "Confidence",
                "How strongly filenames and timestamps supported the automatic proposal. It is a warning signal, " +
                "not proof; weak guesses should be corrected from your knowledge of the files.");

            AppendHeading("Buttons and editable fields");
            AppendDefinition("Move earlier / Move later", SingleLine(ComparisonProjectSetupHelpText.MoveEarlier));
            AppendDefinition("Set master", SingleLine(ComparisonProjectSetupHelpText.SetMaster));
            AppendDefinition("Set baseline", SingleLine(ComparisonProjectSetupHelpText.SetBaseline));
            AppendDefinition("Clear baseline", SingleLine(ComparisonProjectSetupHelpText.ClearBaseline));
            AppendDefinition(
                "Revision label",
                "Changes only the name used in the report. It never renames the source file.");
            AppendDefinition(
                "Parent revision",
                "Assigns the selected revision's direct predecessor. Right-click the revision list for shortcuts " +
                "to use the previous row or clear the parent.");
            AppendDefinition("Browse", SingleLine(ComparisonProjectSetupHelpText.BrowseOutput));
            AppendDefinition(
                "Confirmation checkbox",
                "Unlocks Build comparison after you explicitly accept the visible order, parents, master, and " +
                "optional baseline. Any later edit clears confirmation.");
            AppendDefinition("Build comparison", SingleLine(ComparisonProjectSetupHelpText.BuildComparison));
            AppendDefinition("Cancel", SingleLine(ComparisonProjectSetupHelpText.Cancel));

            AppendHeading("Safety and effective worksheet bounds");
            AppendParagraph(
                "Power Tools hashes and reads each selected revision without opening it for editing. It creates a " +
                "new Review.xlsx and refuses to overwrite an existing destination or a source file.\r\n\r\n" +
                "Cell-content comparison has exactly three channels: stored values, saved calculated values, " +
                "and formulas. Saved calculated values are compared even when the formula also changed. Styles, " +
                "fills, comments, drawings, and package metadata are not comparison changes; they are used only " +
                "for safe reading and integrity checks.\r\n\r\n" +
                "For alignment, each sheet is bounded by its first and last comparison-content rows and columns. " +
                "Physically stored blank or formatting-only cells outside that rectangle are ignored and reported; " +
                "blank-versus-missing storage inside it is retained only for alignment safety. No stored value, " +
                "saved calculated value, or formula is silently trimmed.\r\n\r\n" +
                "Current limits are 2,000,000 physically stored cells per revision and 10,000,000 audited " +
                "across the project; 2,000,000 cells retained inside useful ranges; and 10,000,000 effective " +
                "grid cells. Per-sheet limits are 100,000 rows, 4,096 columns, and 5,000,000 effective grid " +
                "cells. Blank-tail cells are parsed under the input safety limits, reported, and then released " +
                "before alignment.");

            AppendHeading("What happens after Build comparison");
            AppendParagraph(
                "Power Tools validates the lineage, rechecks that every source still matches its inventory hash, " +
                "and then aligns sheets, headers, keys, rows, and columns. If structural evidence is ambiguous, a " +
                "correction window opens instead of guessing. A successful run publishes and verifies a new " +
                "Review.xlsx, then opens it in Excel. Any ignored blank tails are reported after publication.");
        }

        private void AppendHeading(string text)
        {
            if (_content.TextLength > 0) _content.AppendText("\r\n");
            using (var font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold))
                AppendFormatted(text + "\r\n", font, Accent);
        }

        private void AppendParagraph(string text)
        {
            using (var font = new Font("Segoe UI", 9.5F, FontStyle.Regular))
                AppendFormatted(text + "\r\n", font, Ink);
        }

        private void AppendDefinition(string term, string description)
        {
            using (var termFont = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold))
                AppendFormatted(term + ": ", termFont, Ink);
            using (var bodyFont = new Font("Segoe UI", 9.5F, FontStyle.Regular))
                AppendFormatted(description + "\r\n", bodyFont, Ink);
        }

        private void AppendFormatted(string text, Font font, Color color)
        {
            _content.SelectionStart = _content.TextLength;
            _content.SelectionLength = 0;
            _content.SelectionFont = font;
            _content.SelectionColor = color;
            _content.AppendText(text);
        }

        private static string SingleLine(string value) => value.Replace("\r\n", " ");
    }
}
