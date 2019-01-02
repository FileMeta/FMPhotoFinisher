using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace FileMeta
{
    /// <summary>
    /// Supports a mix of traditional keywords and custom metatags
    /// </summary>
    /// <remarks>
    /// <para>A metatag is like a hashtag in that it can be stored wherever text is stored. However,
    /// where a hashtag is only a label or keyword, a metatag is a name and a value.
    /// </para>
    /// <para>Examples:</para>
    /// <para>   &author=Brandt</para>
    /// <para>   &subject=MetaTag_Format</para>
    /// <para>   &date=2018-12-17T21:22:05-06:00</para>
    /// <para>Regular format definition.</para>
    /// <para>A metatag starts with an ampersand - just as a hashtag starts with the hash symbol.
    /// </para>
    /// <para>Next comes the name which follows the same standard as a hashtag - it must be composed
    /// of letters, numbers, and the underscore character. Rigorous implementations should use the
    /// unicode character sets. Specifically Unicode categories: Ll, Lu, Lt, Lo, Lm, Mn, Nd, Pc. For
    /// regular expressiosn this matches the \w chacter class.
    /// </para>
    /// <para>Next is an equals sign.
    /// </para>
    /// <para>Next is the value which is a series of any characters except whitespace or the ampersand.
    /// Whitespace, ampersand, underscore, and the percent character MUST be encoded. A space character
    /// is encoded as the underscore. All other whitespace, ampersand, underscore, or percent characters
    /// are encoded as the percent character followed by two hexadecimal digits. The hexadecimal
    /// the Unicode character code of the encoded character which MUST be in the first 256 characters of
    /// Unicode. Other Unicode characters are given by their literal value. Notably, all characters to be
    /// encoded exist in the first 256.
    /// </para>
    /// <para>The value encoding is deliberately similar to URL query string encoding. However, in
    /// Metatag encoding, the underscore substitutes for a space whereas in URL query strings, the plus
    /// sign substitutes for a space.
    /// </para>
    /// <para>The name IS NOT encoded. Valid names are simply limited to the specified character set.
    /// </para>
    /// <para>This implementation does not handle multiple values for the same key. That's consistent
    /// with the overall Windows Property System.
    /// </para>
    /// </remarks>
    class MetaTagSet
    {
        #region Public Static Members

        /* Matches a metatag which is defined as follows:
           &       An Ampersand
           \w+     One or more "word characters" consisting of the unicode groups for
                   letters numbers, nonspacing marks, numers (decimal digits), punctuation characters (underscore)
           =       The equals sign
           [^\s&]. Zero or more non-whitespace and non-ampersand characters.
        */

        // Matches a metatag that composes the whole string
        static Regex s_rxSingleMetatag = new Regex(
            @"^&(\w+)=([^\s&]*)$",
            RegexOptions.CultureInvariant);

        // Matches metatags that are embedded in a potentially longer string.
        // Includes any leading whitespace in the matched string.
        static Regex s_rxEmbeddedMetatag = new Regex(
            @"&(\w+)=([^\s&]*)",
            RegexOptions.CultureInvariant);

        /// <summary>
        /// Attempt to parse one metatag that composes the whole string
        /// </summary>
        /// <param name="s">The string to parse as a metatag.</param>
        /// <param name="key">The parsed metagag key.</param>
        /// <param name="value">The parsed value.</param>
        /// <returns>True if the string is a valid metatag that was successfully parsed.</returns>
        /// <remarks>To be a valid metatag, the string must start with ampersand and must not have
        /// any embedded whitespace.</remarks>
        public static bool TryParseMetatag(string s, out string key, out string value)
        {
            var match = s_rxSingleMetatag.Match(s);
            if (!match.Success)
            {
                key = null;
                value = null;
                return false;
            }

            key = match.Groups[1].Value;
            value = MetatagDecode(match.Groups[2].Value);
            return true;
        }

        /// <summary>
        /// Parse all metatags embedded in a string.
        /// </summary>
        /// <param name="s">The string that may contain metatags.</param>
        /// <returns>An IEnumerator that will list the metatags as <see cref="KeyValuePair{String, String}"/>.</returns>
        /// <remarks>
        /// <para>Use this method to retrieve a set of metatags embedded in a longer string such as the
        /// comments.
        /// </para>
        /// </remarks>
        public static IEnumerable<KeyValuePair<string, string> > ParseMetatags(string s)
        {
            return new MetatagEnumerable(s);
        }

        /// <summary>
        /// Formats a key and value into a metatag.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="value">The value</param>
        /// <returns>A string containing the properly formatted metatag key and value.</returns>
        public static string MetatagFormat(string key, string value)
        {
            return $"&{key}={MetatagEncode(value)}";
        }

        /// <summary>
        /// Formats a KeyValuePair into a metatag.
        /// </summary>
        /// <param name="pair">A <see cref="KeyValuePair{string, string}"/></param>
        /// <returns>The properly formatted metatag.</returns>
        public static string FormatMetatag(KeyValuePair<string, string> pair)
        {
            return $"&{pair.Key}={MetatagEncode(pair.Value)}";
        }

        /// <summary>
        /// Encode a metatag value
        /// </summary>
        /// <param name="s">The value to encode.</param>
        /// <returns>The encoded value.</returns>
        /// <remarks>
        /// <para>The value portion of a metatag may not contain whitespace or the ampersand. Space characters
        /// are encoded as an underscore. The ampersand, percent, underscore, and all other whitespace characters
        /// are percent encoded similar to URL query strings.
        /// </para>
        /// </remarks>
        public static string MetatagEncode(string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (c == ' ')
                {
                    sb.Append('_');
                }
                else if (Char.IsWhiteSpace(c) || c == '%' || c == '&' || c == '_')
                {
                    int n = (int)c;
                    if (n > 255) n = 32; // In the unexpected case of whitespace outside the ASCII range.
                    sb.Append($"%{n:x2}");
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Decodes the value portion of a metatag.
        /// </summary>
        /// <param name="s">A metatag-encoded string to be decoded.</param>
        /// <returns>The decoded string.</returns>
        /// <remarks>
        /// <para>See <see cref="MetatagEncode"/> for a summary of encoding rules.
        /// </para>
        /// </remarks>
        public static string MetatagDecode(string s)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < s.Length; ++i)
            {
                char c = s[i];
                if (c == '_')
                {
                    sb.Append(' ');
                }
                else if (c == '%' && i < s.Length + 2)
                {
                    int n;
                    if (int.TryParse(s.Substring(i + 1, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture,
                        out n) && n > 0 && n < 256)
                    {
                        sb.Append((char)n);
                        i += 2;
                    }
                    else
                    {
                        sb.Append('%');
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        #endregion Public Static Members

        Dictionary<string, string> m_metaTags = new Dictionary<string, string>();

        #region Properties

        /// <summary>
        /// The metatags stored in a keyword set
        /// </summary>
        public IDictionary<string, string> MetaTags { get { return m_metaTags; } }

        #endregion Properties

        #region Methods

        #endregion Methods

        /// <summary>
        /// Load the metatags from a string - such as a comment.
        /// </summary>
        /// <param name="s">The string from which to load the metatags.</param>
        public void LoadMetatags(string s)
        {
            foreach (var pair in ParseMetatags(s))
            {
                if (string.IsNullOrEmpty(pair.Key)) continue;
                m_metaTags[pair.Key] = pair.Value;
            }
        }

        public string AddMetatagsToString(string s)
        {
            if (s == null) s = string.Empty;

            var tagsEntered = new HashSet<string>();
            var sb = new StringBuilder();
            // Process existing string, suppressing any existing metatags
            // that don't match new values.
            int p = 0;
            foreach(Match match in s_rxEmbeddedMetatag.Matches(s))
            {
                // Transfer any existing text to the stringbuilder
                sb.Append(s, p, match.Index - p);
                p = match.Index;

                // Process the match.
                string key = match.Groups[1].Value;
                string value = MetatagDecode(match.Groups[2].Value);
                string newValue;

                // If a matching tag has already been processed, suppress this instance
                if (tagsEntered.Contains(key))
                {
                    p = match.Index + match.Length;

                    // Trim any leading whitespace (trailing on the stringbuilder)
                    while (sb.Length > 0 && char.IsWhiteSpace(sb[sb.Length - 1]))
                        sb.Length = sb.Length - 1;

                    // If no trailing whitespace, insert one space
                    if (p < s.Length && !char.IsWhiteSpace(s[p]))
                    {
                        sb.Append(' ');
                    }
                }

                // Else, if there's a new value for this metatag. Substitute the new one in-place.
                else if (m_metaTags.TryGetValue(key, out newValue)
                    && !string.Equals(value, newValue))
                {
                    sb.Append(MetatagFormat(key, newValue));
                    p = match.Index + match.Length;
                }

                // Else, preserve the existing value for the metatag verbatim.
                else
                {
                    sb.Append(s, p, match.Length);
                    p += match.Length;
                }

                tagsEntered.Add(key);
            }

            // Transfer over any trailing text in the existing string.
            if (p < s.Length)
            {
                sb.Append(s, p, s.Length - p);
            }

            // Make a list of the remaining tags and sort it.
            var list = new List<KeyValuePair<string, string>>(m_metaTags.Count - tagsEntered.Count);
            foreach(var pair in m_metaTags)
            {
                if (!tagsEntered.Contains(pair.Key))
                {
                    list.Add(pair);
                }
            }
            list.Sort((a, b) => string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase));

            // Insert the remaining metatags
            foreach(var pair in list)
            {
                // Append a space if needed.
                // Metatags can be smashed together - this is for human aesthetics.
                if (sb.Length > 0 && !char.IsWhiteSpace(sb[sb.Length - 1]))
                    sb.Append(' ');

                sb.Append(FormatMetatag(pair));
            }

            return sb.ToString();
        }

        protected class MetatagEnumerable : IEnumerable<KeyValuePair<string, string>>
        {
            string m_s;

            public MetatagEnumerable(string s)
            {
                m_s = s;
            }

            public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
            {
                return new MetatagParser(m_s);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return new MetatagParser(m_s);
            }
        }

        protected class MetatagParser : IEnumerator<KeyValuePair<string, string>>
        {
            MatchCollection m_matches;
            IEnumerator m_enumerator;
            public MetatagParser(string s)
            {
                if (s == null)
                    s = string.Empty;

                m_matches = s_rxEmbeddedMetatag.Matches(s);
                m_enumerator = m_matches.GetEnumerator();
            }

            public KeyValuePair<string, string> Current
            {
                get
                {
                    Match match = m_enumerator.Current as Match;
                    if (match == null) throw new InvalidOperationException();
                    Debug.Assert(match.Success);

                    return new KeyValuePair<string, string>(match.Groups[1].Value, MetatagDecode(match.Groups[2].Value));
                }
            }

            object IEnumerator.Current => throw new NotImplementedException();

            public void Dispose()
            {
                m_enumerator = null;
                m_matches = null;
            }

            public bool MoveNext()
            {
                return m_enumerator.MoveNext();
            }

            public void Reset()
            {
                m_enumerator.Reset();
            }
        }

    }
}
