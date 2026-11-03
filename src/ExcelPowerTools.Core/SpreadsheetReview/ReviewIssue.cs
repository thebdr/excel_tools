using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Text;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public sealed class ReviewIssue
    {
        private ReviewIssue(string code, string message, IDictionary<string, string> context,
            string sourceFile, string sourceMember, int sourceLine)
        {
            if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("An issue code is required.", nameof(code));
            if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("An issue message is required.", nameof(message));
            Code = code; Message = message; IssueId = Guid.NewGuid().ToString("N");
            Context = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(context ?? new Dictionary<string, string>()));
            SourceFile = sourceFile; SourceMember = sourceMember; SourceLine = sourceLine;
        }

        public string Code { get; }
        public string IssueId { get; }
        public string Message { get; }
        public IReadOnlyDictionary<string, string> Context { get; }
        public string SourceFile { get; }
        public string SourceMember { get; }
        public int SourceLine { get; }

        public override string ToString()
        {
            var text = new StringBuilder().Append(Code).Append(" [").Append(IssueId).Append("] ").AppendLine(Message);
            foreach (var item in Context.OrderBy(p => p.Key, StringComparer.Ordinal))
                text.Append("  ").Append(item.Key).Append(": ").AppendLine((item.Value ?? "<null>").Replace("\r", "\\r").Replace("\n", "\\n"));
            return text.Append("  Source: ").Append(SourceFile).Append(':').Append(SourceLine).Append(" (").Append(SourceMember).Append(')').ToString();
        }

        public static ReviewIssue Create(string code, string message, IDictionary<string, string> context = null,
            [CallerFilePath] string sourceFile = "", [CallerMemberName] string sourceMember = "", [CallerLineNumber] int sourceLine = 0) =>
            new ReviewIssue(code, message, context, sourceFile, sourceMember, sourceLine);
    }
}
