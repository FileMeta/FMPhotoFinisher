/*
---
name: MetaTag.cs
description: CodeBit class for extracting and embedding metatags in text fields.
url: https://raw.githubusercontent.com/FileMeta/MetaTag/master/MetaTag.cs
version: 1.0
keywords: CodeBit
dateModified: 2019-01-23
license: https://opensource.org/licenses/BSD-3-Clause
# Metadata in MicroYaml format. See http://filemeta.org/CodeBit.html
...
*/

/*
=== BSD 3 Clause License ===
https://opensource.org/licenses/BSD-3-Clause

Copyright 2019 Brandt Redd

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors
may be used to endorse or promote products derived from this software without
specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace FileMeta
{
    /// <summary>
    /// Static class for parsing and formatting; extracting and embedding metatags
    /// </summary>
    /// <remarks>
    /// <para>A metatag is like a hashtag in that it can be embedded wherever text is stored. However,
    /// where a hashtag is only a label or keyword, a metatag has a name and a value. Thus, metatags
    /// allow custom metadata to be stored in existing text or textual fields such as comments.
    /// </para>
    /// <para>Examples:</para>
    /// <para>   &author=Brandt</para>
    /// <para>   &subject=MetaTag_Format</para>
    /// <para>   &date=2018-12-17T21:22:05-06:00</para>
    /// <para></para>
    /// <para>Format Definition:</para>
    /// <para>A metatag starts with an ampersand - just as a hashtag starts with the hash symbol.
    /// </para>
    /// <para>Next comes the name which follows the same standard as a hashtag - it must be composed
    /// of letters, numbers, and the underscore character. Rigorous implementations should use the
    /// unicode character sets. Specifically Unicode categories: Ll, Lu, Lt, Lo, Lm, Mn, Nd, Pc. For
    /// regular expressions this matches the \w chacter class.
    /// </para>
    /// <para>Next is an equals sign.
    /// </para>
    /// <para>Next is the value which is a series of any characters except the ASCII control range (0x00
    /// to 0x7F), space or the ampersand. Control characters, space, ampersand, underscore, and the percent
    /// character MUST be encoded. A space character is encoded as the underscore. All other control,
    /// ampersand, underscore, or percent characters are encoded as the percent character followed by two
    /// hexadecimal digits. All characters requiring encoding are in the first 256 characters of Unicode, so
    /// two hexadecimal digits are sufficient. Other Unicode characters are given by their literal value.
    /// </para>
    /// <para>The value encoding is deliberately similar to URL query string encoding. However, in
    /// Metatag encoding, the underscore substitutes for a space whereas in URL query strings, the plus
    /// sign substitutes for a space.
    /// </para>
    /// <para>The name IS NOT encoded. Valid names are simply limited to the specified character set.
    /// </para>
    /// </remarks>
    /// <seealso cref="MetaTagSet"/>
    static class MetaTag
    {
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
        public static bool TryParse(string s, out string key, out string value)
        {
            var match = s_rxSingleMetatag.Match(s);
            if (!match.Success)
            {
                key = null;
                value = null;
                return false;
            }

            key = match.Groups[1].Value;
            value = DecodeValue(match.Groups[2].Value);
            return true;
        }

        /// <summary>
        /// Attempt to parse one metatag that composes the whole string
        /// </summary>
        /// <param name="s">The string to parse as a metatag.</param>
        /// <param name="result">A <see cref="KeyValuePair"/> containing the key and value parsed.</param>
        /// <returns>True if the string is a valid metatag that was successfully parsed.</returns>
        /// <remarks>To be a valid metatag, the string must start with ampersand and must not have
        /// any embedded whitespace.</remarks>
        public static bool TryParse(string s, out KeyValuePair<string, string> result)
        {
            string key;
            string value;
            if (!TryParse(s, out key, out value))
            {
                result = new KeyValuePair<string, string>();
                return false;
            }

            result = new KeyValuePair<string, string>(key, value);
            return true;
        }

        /// <summary>
        /// Attempt to parse one metatag that composes the whole string
        /// </summary>
        /// <param name="s">The string to parse as a metatag.</param>
        /// <exception cref="ArgumentException">s is an invalid MetaTag string.</exception>
        /// <returns>A <see cref="KeyValuePair{String, String}"/>containing the result.</returns>
        public static KeyValuePair<string, string> Parse(string s)
        {
            KeyValuePair<string, string> result;
            if (!TryParse(s, out result))
            {
                throw new ArgumentException("Parse failure: Invalid MetaTag String");
            }
            return result;
        }

        /// <summary>
        /// Formats a name and value into a MetaTag string.
        /// </summary>
        /// <param name="name">The key.</param>
        /// <param name="value">The value</param>
        /// <returns>A string containing the properly formatted metatag key and value.</returns>
        public static string Format(string name, string value)
        {
            return $"&{name}={EncodeValue(value)}";
        }

        /// <summary>
        /// Formats a <see cref="KeyValuePair"/> into a metatag.
        /// </summary>
        /// <param name="pair">A <see cref="KeyValuePair{string, string}"/></param>
        /// <returns>The properly formatted metatag.</returns>
        public static string Format(KeyValuePair<string, string> pair)
        {
            return $"&{pair.Key}={EncodeValue(pair.Value)}";
        }

        /// <summary>
        /// Encode a metatag value
        /// </summary>
        /// <param name="s">The value to encode.</param>
        /// <returns>The encoded value.</returns>
        /// <remarks>
        /// <para>The encoded value portion of a metatag may not contain characters lower or equal to 0x1f, the space
        /// character, percent, or the ampersand. Space characters are encoded as an underscore. The ampersand, percent,
        /// underscore, and control characters are percent encoded as in URL query strings.
        /// </para>
        /// </remarks>
        public static string EncodeValue(string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (c == ' ')
                {
                    sb.Append('_');
                }
                else if (c < '\x20' || c == '%' || c == '&' || c == '_')
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
        /// <para>See <see cref="EncodeValue"/> for a summary of encoding rules.
        /// </para>
        /// </remarks>
        public static string DecodeValue(string s)
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
                        out n) && n >= 0 && n < 256)
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

        /// <summary>
        /// Extracts all metatags embedded in a string.
        /// </summary>
        /// <param name="s">A string that may contain metatags.</param>
        /// <returns>An IEnumerator that will list the metatags as <see cref="KeyValuePair{String, String}"/>.</returns>
        /// <remarks>
        /// <para>Use this method to retrieve a set of metatags embedded in a longer string such as the
        /// comments.
        /// </para>
        /// </remarks>
        public static IEnumerable<KeyValuePair<string, string>> Extract(string s)
        {
            return new MetaTagEnumerable(s);
        }

        /// <summary>
        /// Embeds metatags in a string.
        /// </summary>
        /// <param name="s">An existing string, such as comments, to which metatags are added. May be empty or null.</param>
        /// <returns>A string with the metatags added or updated.</returns>
        /// <remarks>
        /// <para>If a metatag already exists in the string that has the same name as one in the set, the value will be
        /// updated in-place (so that the tag remains positioned where it was in the string.
        /// </para>
        /// <para>If a metatag already exists in the string but a corresponding value does not exist in the set, the
        /// existing value is left unchanged.
        /// </para>
        /// <para>If a metatag already exists in the string and a corresponding entry in the set has a null value, the
        /// the existing metatag is removed.
        /// </para>
        /// <para>Metatags in the set that don't already exist in the string are added at the end in alphabetical order unless
        /// the value is null.
        /// </para>
        /// <para>Only supports one value per key. If <paramref name="metaTagSet"/> has multiple values for the same key
        /// the last value will be used.
        /// </para>
        /// </remarks>
        public static string EmbedAndUpdate(string s, IEnumerable<KeyValuePair<string, string>> metaTagSet)
        {
            if (s == null) s = string.Empty;

            var tagsEntered = new HashSet<string>();

            IDictionary<string, string> tagSet = metaTagSet as IDictionary<string, string>;
            if (tagSet == null)
            {
                tagSet = new Dictionary<string, string>();
                foreach (var pair in metaTagSet)
                {
                    tagSet[pair.Key] = pair.Value;
                }
            }

            var sb = new StringBuilder();
            // Process existing string, suppressing any existing metatags
            // that don't have values in the set and updating andy that do have values.
            int p = 0;
            foreach (Match match in s_rxEmbeddedMetatag.Matches(s))
            {
                // Transfer any existing text to the stringbuilder
                sb.Append(s, p, match.Index - p);
                p = match.Index;

                // Process the match.
                string key = match.Groups[1].Value;
                string value = DecodeValue(match.Groups[2].Value);
                string newValue;
                bool inSet = tagSet.TryGetValue(key, out newValue);

                // If a tag with the same value has already been processed,
                // or if the new value is null, suppress this tag.
                if (tagsEntered.Contains(key) || (inSet && newValue == null))
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
                else if (inSet && !string.Equals(value, newValue, StringComparison.Ordinal))
                {
                    sb.Append(Format(key, newValue));
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
            var list = new List<KeyValuePair<string, string>>(tagSet.Count - tagsEntered.Count);
            foreach (var pair in tagSet)
            {
                if (!tagsEntered.Contains(pair.Key))
                {
                    list.Add(pair);
                }
            }
            list.Sort((a, b) => string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase));

            // Insert the remaining metatags
            foreach (var pair in list)
            {
                // Append a space if needed.
                // Metatags can be smashed together - this is for human aesthetics.
                if (sb.Length > 0 && !char.IsWhiteSpace(sb[sb.Length - 1]))
                    sb.Append(' ');

                sb.Append(Format(pair));
            }

            return sb.ToString();
        }

        protected class MetaTagEnumerable : IEnumerable<KeyValuePair<string, string>>
        {
            string m_s;

            public MetaTagEnumerable(string s)
            {
                m_s = s;
            }

            public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
            {
                return new MetaTagEnumerator(m_s);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return new MetaTagEnumerator(m_s);
            }
        }

        protected class MetaTagEnumerator : IEnumerator<KeyValuePair<string, string>>
        {
            MatchCollection m_matches;
            IEnumerator m_enumerator;
            public MetaTagEnumerator(string s)
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

                    return new KeyValuePair<string, string>(match.Groups[1].Value, DecodeValue(match.Groups[2].Value));
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

    } // Class MetaTag
    
    /// <summary>
    /// Represents a set of <see cref="MetaTag"/> - typically embedded in text such as a comment field.
    /// </summary>
    /// <remarks>
    /// <para>This implementation expects only one value for each key. That's consistent
    /// with the Windows Property System.
    /// </para>
    /// </remarks>
    /// <seealso cref="MetaTag"/>
    class MetaTagSet : Dictionary<string, string>
    {
        /// <summary>
        /// Load the metatags from a string - such as a comment.
        /// </summary>
        /// <param name="s">The string from which to load the metatags.</param>
        /// <seealso cref="MetaTag.Extract(string)"/>
        public void Load(string s)
        {
            foreach (var pair in MetaTag.Extract(s))
            {
                this[pair.Key] = pair.Value;
            }
        }

        /// <summary>
        /// Update and embed metatags into an existing string - such as a comment.
        /// </summary>
        /// <param name="s">The string into which to embed the metatags.</param>
        /// <returns>A string with updated metatags.</returns>
        /// <seealso cref="MetaTag.EmbedAndUpdateMetaTags(string, IEnumerable{KeyValuePair{string, string}})"/>
        public string EmbedAndUpdate(string s)
        {
            return MetaTag.EmbedAndUpdate(s, this);
        }
    }
}
