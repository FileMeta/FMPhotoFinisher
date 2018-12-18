using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

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
    /// <para>A metatag starts with an ampersand - just as a hashtag starts with the hash symbol.</para>
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
        static Regex s_rxMetatag = new Regex(
            @"^&(\w+)=([^\s&]*)$",
            RegexOptions.CultureInvariant);

        public static bool TryParseMetatag(string s, out string key, out string value)
        {
            var match = s_rxMetatag.Match(s);
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

        public static string FormatMetatag(string key, string value)
        {
            return $"&{key}={MetatagEncode(value)}";
        }

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

        HashSet<string> m_keywords = new HashSet<string>();
        Dictionary<string, string> m_metaTags = new Dictionary<string, string>();

        #region Properties

        /// <summary>
        /// Conventional keywords in the keyword set
        /// </summary>
        public ISet<string> Keywords { get { return m_keywords; } }
        
        /// <summary>
        /// The metatags stored in a keyword set
        /// </summary>
        public IDictionary<string, string> MetaTags { get { return m_metaTags; } }

        #endregion Properties

        #region Methods

        #endregion Methods

        public void LoadKeywords(string[] keywords)
        {
            if (keywords != null)
            {
                foreach (string s in keywords)
                {
                    if (string.IsNullOrEmpty(s)) continue;
                    string key, value;
                    if (TryParseMetatag(s, out key, out value))
                    {
                        m_metaTags[key] = value;
                    }
                    else
                    {
                        m_keywords.Add(s);
                    }
                }
            }
        }

        public string[] ToKeywords()
        {
            var list = new List<string>(m_keywords.Count + m_metaTags.Count);

            // Load the keywords into a list
            foreach (var k in m_keywords)
            {
                if (!string.IsNullOrEmpty(k))
                {
                    list.Add(k);
                }
            }

            // Add the metatags
            foreach (var pair in m_metaTags)
            {
                if (!string.IsNullOrEmpty(pair.Key))
                {
                    list.Add(FormatMetatag(pair.Key, pair.Value ?? string.Empty));
                }
            }

            // Sort
            list.Sort(delegate (string a, string b)
            {
                // Sort metatags after keywords. We already verified that all
                // keys have at least one character.
                if (a[0] != b[0])
                {
                    if (a[0] == '&') return 1;
                    if (b[0] == '&') return -1;
                }
                return string.CompareOrdinal(a, b);
            });

            return list.ToArray();
        }

    }
}
