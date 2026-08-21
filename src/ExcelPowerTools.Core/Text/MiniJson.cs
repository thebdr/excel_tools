using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core.Text
{
    public enum JsonKind { Object, Array, String, Number, Bool, Null }

    /// <summary>
    /// A parsed JSON value. Numbers keep their raw token text (culture-safe; the caller parses),
    /// objects preserve property order. Accessors never throw: a missing or differently-typed
    /// property yields the fallback, so readers tolerate unknown structure by construction.
    /// </summary>
    public sealed class JsonValue
    {
        public JsonKind Kind;
        public List<KeyValuePair<string, JsonValue>> Properties;   // Kind == Object
        public List<JsonValue> Items;                              // Kind == Array
        public string Text;                                        // String value, or raw Number token
        public bool Flag;                                          // Kind == Bool

        public JsonValue Get(string name)
        {
            if (Kind != JsonKind.Object || Properties == null) return null;
            foreach (var p in Properties)
                if (p.Key == name) return p.Value;
            return null;
        }

        public string Str(string name, string fallback = "")
        {
            JsonValue v = Get(name);
            return v != null && (v.Kind == JsonKind.String || v.Kind == JsonKind.Number) ? v.Text : fallback;
        }

        public bool BoolOf(string name, bool fallback)
        {
            JsonValue v = Get(name);
            return v != null && v.Kind == JsonKind.Bool ? v.Flag : fallback;
        }

        public List<JsonValue> Arr(string name)
        {
            JsonValue v = Get(name);
            return v != null && v.Kind == JsonKind.Array && v.Items != null ? v.Items : new List<JsonValue>();
        }
    }

    /// <summary>
    /// Minimal, dependency-free JSON reader (the repo's serializers are write-only). Strict JSON:
    /// full escape handling including \uXXXX (surrogate halves recombine naturally), no comments,
    /// no trailing commas. Errors carry the character position.
    /// </summary>
    public static class MiniJson
    {
        public static bool TryParse(string json, out JsonValue value, out string error)
        {
            value = null;
            error = null;
            if (json == null) { error = "No JSON text."; return false; }

            int i = 0;
            if (!ParseValue(json, ref i, 0, out value, out error)) return false;

            SkipWhitespace(json, ref i);
            if (i != json.Length)
            {
                error = "Unexpected content after the JSON value at position " + i + ".";
                value = null;
                return false;
            }
            return true;
        }

        private const int MaxDepth = 64;

        private static bool ParseValue(string s, ref int i, int depth, out JsonValue value, out string error)
        {
            value = null;
            error = null;
            if (depth > MaxDepth) { error = "JSON nesting is too deep."; return false; }

            SkipWhitespace(s, ref i);
            if (i >= s.Length) { error = "Unexpected end of JSON."; return false; }

            char c = s[i];
            switch (c)
            {
                case '{': return ParseObject(s, ref i, depth, out value, out error);
                case '[': return ParseArray(s, ref i, depth, out value, out error);
                case '"':
                    if (!ParseString(s, ref i, out string str, out error)) return false;
                    value = new JsonValue { Kind = JsonKind.String, Text = str };
                    return true;
                case 't':
                    if (!ExpectLiteral(s, ref i, "true", out error)) return false;
                    value = new JsonValue { Kind = JsonKind.Bool, Flag = true };
                    return true;
                case 'f':
                    if (!ExpectLiteral(s, ref i, "false", out error)) return false;
                    value = new JsonValue { Kind = JsonKind.Bool, Flag = false };
                    return true;
                case 'n':
                    if (!ExpectLiteral(s, ref i, "null", out error)) return false;
                    value = new JsonValue { Kind = JsonKind.Null };
                    return true;
                default:
                    return ParseNumber(s, ref i, out value, out error);
            }
        }

        private static bool ParseObject(string s, ref int i, int depth, out JsonValue value, out string error)
        {
            value = null;
            i++;   // '{'
            var props = new List<KeyValuePair<string, JsonValue>>();

            SkipWhitespace(s, ref i);
            if (i < s.Length && s[i] == '}')
            {
                i++;
                value = new JsonValue { Kind = JsonKind.Object, Properties = props };
                error = null;
                return true;
            }

            while (true)
            {
                SkipWhitespace(s, ref i);
                if (i >= s.Length || s[i] != '"') { error = "Expected a property name at position " + i + "."; return false; }
                if (!ParseString(s, ref i, out string name, out error)) return false;

                SkipWhitespace(s, ref i);
                if (i >= s.Length || s[i] != ':') { error = "Expected ':' at position " + i + "."; return false; }
                i++;

                if (!ParseValue(s, ref i, depth + 1, out JsonValue member, out error)) return false;
                props.Add(new KeyValuePair<string, JsonValue>(name, member));

                SkipWhitespace(s, ref i);
                if (i >= s.Length) { error = "Unterminated object."; return false; }
                if (s[i] == ',') { i++; continue; }
                if (s[i] == '}') { i++; break; }
                error = "Expected ',' or '}' at position " + i + ".";
                return false;
            }

            value = new JsonValue { Kind = JsonKind.Object, Properties = props };
            return true;
        }

        private static bool ParseArray(string s, ref int i, int depth, out JsonValue value, out string error)
        {
            value = null;
            i++;   // '['
            var items = new List<JsonValue>();

            SkipWhitespace(s, ref i);
            if (i < s.Length && s[i] == ']')
            {
                i++;
                value = new JsonValue { Kind = JsonKind.Array, Items = items };
                error = null;
                return true;
            }

            while (true)
            {
                if (!ParseValue(s, ref i, depth + 1, out JsonValue item, out error)) return false;
                items.Add(item);

                SkipWhitespace(s, ref i);
                if (i >= s.Length) { error = "Unterminated array."; return false; }
                if (s[i] == ',') { i++; continue; }
                if (s[i] == ']') { i++; break; }
                error = "Expected ',' or ']' at position " + i + ".";
                return false;
            }

            value = new JsonValue { Kind = JsonKind.Array, Items = items };
            return true;
        }

        private static bool ParseString(string s, ref int i, out string result, out string error)
        {
            result = null;
            error = null;
            i++;   // opening quote
            var sb = new StringBuilder();

            while (true)
            {
                if (i >= s.Length) { error = "Unterminated string."; return false; }
                char c = s[i++];
                if (c == '"') break;
                if (c == '\\')
                {
                    if (i >= s.Length) { error = "Unterminated escape."; return false; }
                    char e = s[i++];
                    switch (e)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (i + 4 > s.Length ||
                                !ushort.TryParse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort code))
                            {
                                error = "Invalid \\u escape at position " + i + ".";
                                return false;
                            }
                            // Surrogate halves are appended as-is; consecutive escapes recombine
                            // into the pair because .NET strings are UTF-16.
                            sb.Append((char)code);
                            i += 4;
                            break;
                        default:
                            error = "Invalid escape '\\" + e + "' at position " + (i - 1) + ".";
                            return false;
                    }
                }
                else if (c < 0x20)
                {
                    error = "Unescaped control character at position " + (i - 1) + ".";
                    return false;
                }
                else
                {
                    sb.Append(c);
                }
            }

            result = sb.ToString();
            return true;
        }

        private static bool ParseNumber(string s, ref int i, out JsonValue value, out string error)
        {
            value = null;
            error = null;
            int start = i;

            if (i < s.Length && s[i] == '-') i++;
            if (i >= s.Length || s[i] < '0' || s[i] > '9')
            {
                error = "Unexpected character at position " + start + ".";
                return false;
            }
            // JSON forbids leading zeros: "0" stands alone, anything else must not start with '0'.
            if (s[i] == '0') i++;
            else while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i < s.Length && s[i] == '.')
            {
                i++;
                if (i >= s.Length || s[i] < '0' || s[i] > '9') { error = "Malformed number at position " + start + "."; return false; }
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            }
            if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
            {
                i++;
                if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
                if (i >= s.Length || s[i] < '0' || s[i] > '9') { error = "Malformed number at position " + start + "."; return false; }
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            }

            value = new JsonValue { Kind = JsonKind.Number, Text = s.Substring(start, i - start) };
            return true;
        }

        private static bool ExpectLiteral(string s, ref int i, string literal, out string error)
        {
            error = null;
            if (i + literal.Length <= s.Length && string.CompareOrdinal(s, i, literal, 0, literal.Length) == 0)
            {
                i += literal.Length;
                return true;
            }
            error = "Unexpected token at position " + i + ".";
            return false;
        }

        private static void SkipWhitespace(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
        }
    }

    /// <summary>
    /// Append-style JSON writer helpers, mirroring the private ones of the comparison module's
    /// serializer so Core code can emit JSON without depending on the add-in assembly.
    /// </summary>
    public static class MiniJsonWriter
    {
        /// <summary>
        /// Reformats compact JSON with newlines and two-space indentation — the sidecars are
        /// read by humans too. Purely structural (string-content aware); the parser ignores
        /// the added whitespace, so round-trips are untouched.
        /// </summary>
        public static string Indent(string json)
        {
            if (string.IsNullOrEmpty(json)) return json ?? "";
            var sb = new StringBuilder(json.Length * 2);
            int depth = 0;
            bool inString = false;

            void NewLine()
            {
                sb.Append('\n');
                sb.Append(' ', depth * 2);
            }

            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];
                if (inString)
                {
                    sb.Append(c);
                    if (c == '\\' && i + 1 < json.Length) sb.Append(json[++i]);
                    else if (c == '"') inString = false;
                    continue;
                }
                switch (c)
                {
                    case '"':
                        inString = true;
                        sb.Append(c);
                        break;
                    case '{':
                    case '[':
                        sb.Append(c);
                        if (i + 1 < json.Length && (json[i + 1] == '}' || json[i + 1] == ']'))
                        {
                            sb.Append(json[++i]);   // {} and [] stay on one line
                        }
                        else
                        {
                            depth++;
                            NewLine();
                        }
                        break;
                    case '}':
                    case ']':
                        depth--;
                        NewLine();
                        sb.Append(c);
                        break;
                    case ',':
                        sb.Append(c);
                        NewLine();
                        break;
                    case ':':
                        sb.Append(": ");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        public static void Property(StringBuilder json, string name, string value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':');
            AppendString(json, value ?? "");
        }

        public static void BoolProperty(StringBuilder json, string name, bool value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(value ? ":true" : ":false");
        }

        public static void NumberProperty(StringBuilder json, string name, long value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':').Append(value.ToString(CultureInfo.InvariantCulture));
        }

        public static void AppendString(StringBuilder json, string value)
        {
            json.Append('"');
            value = value ?? "";
            for (int index = 0; index < value.Length; index++)
            {
                char character = value[index];
                switch (character)
                {
                    case '"': json.Append("\\\""); break;
                    case '\\': json.Append("\\\\"); break;
                    case '\b': json.Append("\\b"); break;
                    case '\f': json.Append("\\f"); break;
                    case '\n': json.Append("\\n"); break;
                    case '\r': json.Append("\\r"); break;
                    case '\t': json.Append("\\t"); break;
                    default:
                        if (character < 0x20 || character == '\u2028' || character == '\u2029' ||
                            char.IsSurrogate(character))
                            json.Append("\\u").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
                        else json.Append(character);
                        break;
                }
            }
            json.Append('"');
        }
    }
}
