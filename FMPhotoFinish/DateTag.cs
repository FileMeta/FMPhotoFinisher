/*
---
name: DateTag.cs
description: CodeBit class that represents a Date metadata tag including DateTime, TimeZone, and Precision components. Includes parsing and formatting methods.
url: https://raw.githubusercontent.com/FileMeta/DateTag/master/DateTag.cs
version: 1.2
keywords: CodeBit
dateModified: 2019-01-30
license: https://opensource.org/licenses/BSD-3-Clause
dependsOn: https://raw.githubusercontent.com/FileMeta/TimeZoneTag/master/TimeZoneTag.cs
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
using System.Text;
using System.Globalization;

namespace FileMeta
{
    /// <summary>
    /// Represents a Date metadata tag.
    /// </summary>
    /// <remarks>
    /// <para>A proper date metadata tag has three components. 1) the date and time the event
    /// occurred (in local time) 2) the timezone of the event and 3) the precision of the date
    /// and time.
    /// </para>
    /// <para>This class includes properties for the three components and methods for parsing
    /// and formatting the tag according to <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF</see>
    /// specifications. Notably, the W3CDTF format is a single string that includes all three
    /// components (dateTime, timezone, and precision. Lower precision date-time values are
    /// represented by only including the parts that are significant (year, month, day, hour, etc.).
    /// </para>
    /// <para>Precision is represented in terms of significant digits: 4 = year, 6 = month,
    /// 8 = day, 10 = hour, 12 = minute, 14 = second, 17 = millisecond, 20 = microsecond,
    /// 21 = tick (100 nanoseconds).
    /// </para>
    /// <para>DateTag is immutable.
    /// </para>
    /// </remarks>
    class DateTag
    {

        #region Constants

        public const int PrecisionMin = 4;
        public const int PrecisionYear = 4;
        public const int PrecisionMonth = 6;
        public const int PrecisionDay = 8;
        public const int PrecisionHour = 10;
        public const int PrecisionMinute = 12;
        public const int PrecisionSecond = 14;
        public const int PrecisionMillisecond = 17;
        public const int PrecisionMicrosecond = 20;
        public const int PrecisionTick = 21;
        public const int PrecisionMax = 21;
        public static readonly DateTime ZeroDate = new DateTime(0L);

        const long c_ticksPerSecond = 10000000;
        const long c_ticksPerMillisecond = 10000;
        const long c_ticksPerMicrosecond = 10;

        #endregion Constants

        #region Static Methods

        /// <summary>
        /// Parses a metadata date tag into a <see cref="DateTag"/> including <see cref="DateTime"/>, <see cref="TimeZoneTag"/>, and significant digits.
        /// </summary>
        /// <param name="dateTag">The value to be parsed in <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF</see> format.</param>
        /// <param name="result">The result of the parsing.</param>
        /// <returns>True if successful, else false.</returns>
        /// <remarks>
        /// <para>The <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF</see> format has date and timezone portions.
        /// This method parses both.</para>
        /// <para>If the timezone portion is not included in the input string then the resulting <paramref name="timezone"/>
        /// will have <see cref="Kind"/> set to <see cref="TimeZoneKind.ForceLocal"/>.
        /// </para>
        /// <para>If the timezone portion is set to "Z" indicating UTC, then the resulting <paramref name="timezone"/>
        /// will have <see cref="Kind"/> set to <see cref="TimeZoneKind.ForceUtc"/> and the <see cref="UtcOffset"/>
        /// will be zero.
        /// </para>
        /// <para>The W2CDTF format permits partial date-time values. For example "2018" is just a year with no
        /// other information. The <paramref name="precision"/> value indicates how much detail is included
        /// as follows: 4 = year, 6 = month, 8 = day, 10 = hour, 12 = minute, 14 = second, 17 = millisecond, 20 = microsecond,
        /// 21 = tick (100 nanoseconds).
        /// </para>
        /// </remarks>
        public static bool TryParse(string dateTag, out DateTag result)
        {
            // Init values for failure case
            result = new DateTag(ZeroDate, TimeZoneTag.Zero, 0);

            // Init parts
            int year = 0;
            int month = 1;
            int day = 1;
            int hour = 12;  // Noon
            int minute = 0;
            int second = 0;
            long ticks = 0;

            // Track position
            int pos = 0;

            if (dateTag.Length < 4) return false;

            if (!int.TryParse(dateTag.Substring(0, 4), out year)
                || year < 1 || year > 9999) return false;
            int precision = PrecisionYear;
            pos = 4;
            if (dateTag.Length > 5 && dateTag[4] == '-')
            {
                if (!int.TryParse(dateTag.Substring(5, 2), out month)
                    || month < 1 || month > 12) return false;
                precision = PrecisionMonth;
                pos = 7;
                if (dateTag.Length > 8 && dateTag[7] == '-')
                {
                    if (!int.TryParse(dateTag.Substring(8, 2), out day)
                        || day < 1 || day > DateTime.DaysInMonth(year, month)) return false;
                    precision = PrecisionDay;
                    pos = 10;
                    if (dateTag.Length > 11 && (dateTag[10] == 'T' || dateTag[10] == ' ')) // Even though W3CDTF and ISO 8601 specify 'T' separating date and time, tolerate a space as an alternative.
                    {
                        if (!int.TryParse(dateTag.Substring(11, 2), out hour)
                            || hour < 0 || hour > 23) return false;
                        precision = PrecisionHour;
                        pos = 13;
                        if (dateTag.Length > 14 && dateTag[13] == ':')
                        {
                            if (!int.TryParse(dateTag.Substring(14, 2), out minute)
                                || minute < 0 || minute > 59) return false;
                            precision = PrecisionMinute;
                            pos = 16;
                            if (dateTag.Length > 17 && dateTag[16] == ':')
                            {
                                if (!int.TryParse(dateTag.Substring(17, 2), out second)
                                    || second < 0 || second > 59) return false;
                                precision = PrecisionSecond;
                                pos = 19;
                                if (dateTag.Length > 20 && dateTag[19] == '.')
                                {
                                    ++pos;
                                    int anchor = pos;
                                    while (pos < dateTag.Length && char.IsDigit(dateTag[pos]))
                                        ++pos;

                                    precision = PrecisionSecond + (pos - anchor);
                                    if (precision > PrecisionMax)
                                        precision = PrecisionMax;

                                    double d;
                                    if (!double.TryParse(dateTag.Substring(anchor, pos - anchor), out d)) return false;
                                    ticks = (long)(d * Math.Pow(10.0, 7.0 - (pos - anchor)));
                                }
                            }
                        }
                    }
                }
            }

            // Attempt to parse the timezone
            TimeZoneTag timezone;
            DateTimeKind dtk = DateTimeKind.Unspecified;
            if (pos < dateTag.Length)
            {
                if (!TimeZoneTag.TryParse(dateTag.Substring(pos), out timezone)) return false;
                dtk = (timezone.Kind == TimeZoneKind.ForceUtc) ? DateTimeKind.Utc : DateTimeKind.Local;
            }
            else
            {
                timezone = TimeZoneTag.ForceLocal;
                dtk = DateTimeKind.Local;
            }

            result = new DateTag(new DateTime(year, month, day, hour, minute, second, dtk).AddTicks(ticks),
                timezone, precision);
            return true;
        }

        /// <summary>
        /// Parses a metadata date tag into <see cref="DateTime"/>, <see cref="TimeZoneTag"/>, and significant digits.
        /// </summary>
        /// <param name="dateTag">The value to be parsed in <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF</see> format.</param>
        /// <returns>The resulting <see cref="DateTag"/>.</returns>
        /// <exception cref="ArgumentException">The dateTag was not in a supported format.</exception>
        /// <seealso cref="TryParse(string, out DateTag)"/>
        public static DateTag Parse(string dateTag)
        {
            DateTag value;
            if (!TryParse(dateTag, out value))
            {
                throw new ArgumentException();
            }
            return value;
        }

        /// <summary>
        /// Detects sub-second precision from a <see cref="DateTime"/> value;
        /// </summary>
        /// <param name="dt">The value on which to detect precision.</param>
        /// <returns>A precision value betweein <see cref="PrecisionSecond"/> and <see cref="PrecisionTick"/> inclusive.</returns>
        /// <remarks>
        /// <para>Detects the sub-second precision based on number of zero digits after the decimal
        /// point. Value will be <see cref="PrecisionSecond"/> (14), <see cref="PrecisionMillisecond"/> (18),
        /// <see cref="PrecisionMicrosecond"/> (20), or <see cref="PrecisionTick"/> (21).
        /// </para>
        /// </remarks>
        public static int DetectPrecision(DateTime dt)
        {
            if (dt.Ticks % c_ticksPerSecond == 0L)
                return PrecisionSecond;
            if (dt.Ticks % c_ticksPerMillisecond == 0L)
                return PrecisionMillisecond;
            if (dt.Ticks % c_ticksPerMicrosecond == 0L)
                return PrecisionMicrosecond;
            return PrecisionTick;
        }

        /// <summary>
        /// Adapts a custom date-time format string to the specified precision
        /// </summary>
        /// <param name="datePrecision">The date precision between 4 and 21.</param>
        /// <param name="cultureInfo">The <see cref="CultureInfo"/> associated with the format to be displayed.</param>
        /// <param name="customFormat">A custom Date-Time format string.</param>
        /// <returns>A custom date-time format string that only includes the elements appropriate for the specified precision.</returns>
        /// <remarks>
        /// <para>The input string should be a custom date-time format string that contains all of the
        /// elements that the application may require. For example, the default U.S. english format is
        /// "dddd, MMMM dd, yyyy h:mm:ss tt". While preserving order, this method will remove elements and
        /// associated delimiters from the string that are beyond the specified precision level.
        /// </para>
        /// </remarks>
        public static string PrecisionAdaptFormatString(int datePrecision, string customFormat)
        {
            // Strip out all percent characters. They are most likely unnecessary.
            // We'll add one back in at the end if it turns out to be important.
            customFormat = customFormat.Replace("%", null);

            if (datePrecision > PrecisionSecond)
            {
                customFormat = LimitFormatComponent(customFormat, 'F', datePrecision - PrecisionSecond);
                customFormat = LimitFormatComponent(customFormat, 'f', datePrecision - PrecisionSecond);

                // Obscure case where we produced a format string with just formatting character, have to prefix with '%'
                if (customFormat.Length == 1 && IsDateFormatChar(customFormat[0]))
                    customFormat = string.Concat("%", customFormat);

                return customFormat;          
            }

            if (datePrecision <= PrecisionSecond)
            {
                customFormat = RemoveFormatComponent(customFormat, 'F');
                customFormat = RemoveFormatComponent(customFormat, 'f');
            }
            if (datePrecision < PrecisionSecond)
            {
                customFormat = RemoveFormatComponent(customFormat, 's');
            }
            if (datePrecision < PrecisionMinute)
            {
                customFormat = RemoveFormatComponent(customFormat, 'm');
            }
            if (datePrecision < PrecisionHour)
            {
                customFormat = RemoveFormatComponent(customFormat, 'H'); // 24-hour format
                customFormat = RemoveFormatComponent(customFormat, 'h'); // 12-hour format
                customFormat = RemoveFormatComponent(customFormat, 'K'); // Timezone info
                customFormat = RemoveFormatComponent(customFormat, 'z'); // Timezone offset
                customFormat = RemoveFormatComponent(customFormat, 't'); // AM/PM
            }
            if (datePrecision < PrecisionDay)
            {
                customFormat = RemoveFormatComponent(customFormat, 'd'); // Captures both day of month and day of week
            }
            if (datePrecision < PrecisionMonth)
            {
                customFormat = RemoveFormatComponent(customFormat, 'M'); // Captures both day of month and day of week
            }
            // For precisions lower than month or year, still include the year.

            // Obscure case where we produced a format string with just formatting character, have to prefix with '%'
            if (customFormat.Length == 1 && IsDateFormatChar(customFormat[0]))
                customFormat = string.Concat("%", customFormat);

            return customFormat;
        }

        // Support for PrecisionAdaptFormatString
        // Remove a formatting component and preceding or succeeding literals
        private static string RemoveFormatComponent(string format, char compChar)
        {
            int preLiteral = int.MaxValue; // Start of preceding literal segment.
            int i = 0;

            while (i < format.Length)
            {
                // if the component was found, remove it.
                if (format[i] == compChar)
                {
                    int compStart = i;
                    do { ++i; } while (i < format.Length && format[i] == compChar);

                    // Special case for commas following days (strange American formatting)
                    if (compChar == 'd' && i < format.Length && format[i] == ',')
                        ++i;

                    // If we found a leading delimiter, remove it and the component
                    if (preLiteral < compStart)
                    {
                        format = format.Remove(preLiteral, i - preLiteral);
                        i = preLiteral;
                    }

                    // Else, remove the component plus trailing delimiters
                    else
                    {
                        i = SkipLiterals(format, i);
                        format = format.Remove(compStart, i - compStart);
                        i = compStart;
                    }
                    preLiteral = int.MaxValue;
                }

                // If some other component, skip it
                else if (IsDateFormatChar(format[i]))
                {
                    char c = format[i];
                    do { ++i; } while (i < format.Length && format[i] == c);
                    preLiteral = int.MaxValue;
                }

                // Else, skip the literal sequence
                else
                {
                    preLiteral = i;
                    i = SkipLiterals(format, i);
                }
            }

            return format;
        }

        // Support for PrecisionAdaptFormatString
        // Indicate whether this is a date formatting character (otherwise it's a literal)
        private static bool IsDateFormatChar(char c)
        {
            // This is the most efficient way to code it because the compiler implements
            // a fast selection algorithm.
            switch (c)
            {
                case 'F': return true;
                case 'H': return true;
                case 'K': return true;
                case 'M': return true;
                case 'd': return true;
                case 'f': return true;
                case 'g': return true;
                case 'h': return true;
                case 'm': return true;
                case 's': return true;
                case 't': return true;
                case 'y': return true;
                case 'z': return true;
                default: return false;
            }
        }

        // Support for PrecisionAdaptFormatString
        // Limit a format component (always fractions of a second) to the specified number of characters.
        // So, for example, the fraction may be limited to 3 digits (milliseconds)
        private static string LimitFormatComponent(string format, char compChar, int maxCount)
        {
            System.Diagnostics.Debug.Assert(maxCount > 0);
            int i = 0;

            while (i < format.Length)
            {
                if (format[i] == compChar)
                {
                    int compStart = i;
                    do { ++i; } while (i < format.Length && format[i] == compChar);
                    if (i-compStart > maxCount)
                    {
                        format = format.Remove(compStart, (i - compStart) - maxCount);
                        i = compStart + maxCount;
                    }
                }
                else if (IsDateFormatChar(format[i]))
                {
                    ++i;
                }
                else
                {
                    i = SkipLiterals(format, i);
                }
            }

            return format;
        }

        // Support for PrecisionAdaptFormatString
        // Skip a sequence of literals in a format string
        private static int SkipLiterals(string format, int pos)
        {
            int end = format.Length;
            while (pos < end && !IsDateFormatChar(format[pos]))
            {
                if (format[pos] == '\\')
                {
                    ++pos; // skip the following character
                }
                else if (format[pos] == '"')
                {
                    ++pos;
                    while (pos < end && format[pos] != '"')
                        ++pos;
                }
                else if (format[pos] == '\'')
                {
                    ++pos;
                    while (pos < end && format[pos] != '\'')
                        ++pos;
                }
                if (pos < end) ++pos; // Skip one more character
            }
            return pos;
        }


        #endregion

        #region Member Variables

        long m_dateTicks;

        #endregion


        #region Public Properties

        /// <summary>
        /// The date value in ticks.
        /// </summary>
        /// <seealso cref="DateTime.Ticks"/>
        public long Ticks { get { return m_dateTicks; } }

        /// <summary>
        /// The <see cref="DateTime"/> value. Always uses <see cref="DateTimeKind.Local"/> even if TimeZoneTag is ForceUtc.
        /// </summary>
        public DateTime Date { get { return new DateTime(m_dateTicks, DateTimeKind.Local); } }

        /// <summary>
        /// The <see cref="DateTime"/> value in UTC. Always uses <see cref="DateTimeKind.Utc"/> even if TimeZoneTag is ForceLocal.
        /// </summary>
        public DateTime DateUtc { get { return TimeZone.ToUtc(Date); } }

        /// <summary>
        /// The timezone value. Not relevant if precision is less than 10.
        /// </summary>
        public TimeZoneTag TimeZone { get; private set; }

        /// <summary>
        /// The precision of the <see cref="Date"/> in significant digits.
        /// </summary>
        /// <remarks>
        /// Significant digits: 4 = year, 6 = month, 8 = day, 10 = hour,
        /// 12 = minute, 14 = second, 17 = millisecond, 20 = microsecond,
        /// 21 = tick (100 nanoseconds).
        /// </para>
        /// </remarks>
        public int Precision { get; private set; }

        #endregion Public Properties

        #region Constructor and Methods

        /// <summary>
        /// Constructs a DateTag from constituent values
        /// </summary>
        /// <param name="date">A <see cref="DateTime"/> value.</param>
        /// <param name="timeZone">A <see cref="TimeZoneTag"/> value or null if unknown.</param>
        /// <param name="precision">Precision in terms of significant digits. If zero
        /// then set to maximum (<see cref="PrecisionMax"/>).</param>
        /// <remarks>
        /// <para>If timeZone is null, the timezone will be set to <see cref="TimeZoneTag.ForceLocal"/>
        /// if the <paramref name="date"/> <see cref="DateTime.Kind"/> is <see cref="DateTimeKind.Local"/>
        /// or <see cref="TimeZoneTag.Unknown"/>, and to <see cref="TimeZoneTag.ForceUtc"/> if
        /// <see cref="DateTime.Kind"/> is <see cref="DateTimeKind.Utc"/>.
        /// </para>
        /// <para>If precision is zero, the precision is detected by the number of trailing zeros
        /// after the seconds decimal point. The lowest precision detected is <see cref="PrecisionSecond"/>.
        /// See <see cref="DetectPrecision(DateTime)"/>.
        /// </para>
        /// </remarks>
        public DateTag(DateTime date, TimeZoneTag timeZone = null, int precision = 0)
        {
            // Default the timezone value if needed.
            if (timeZone == null)
            {
                switch (date.Kind)
                {
                    case DateTimeKind.Local:
                        timeZone = TimeZoneTag.ForceLocal;
                        break;

                    case DateTimeKind.Utc:
                        timeZone = TimeZoneTag.ForceUtc;
                        break;

                    default:
                        timeZone = TimeZoneTag.Zero;
                        break;
                }
            }

            // Change date to a local timezone if needed
            if (date.Kind == DateTimeKind.Utc)
            {
                date = timeZone.ToLocal(date);
            }

            // Limit precision to compatible range
            if (precision > PrecisionMax) precision = PrecisionMax;
            if (precision < PrecisionMin) precision = DetectPrecision(date);

            m_dateTicks = date.Ticks;
            TimeZone = timeZone;
            Precision = precision;
        }

        /// <summary>
        /// Constructs a DateTag from constituent values
        /// </summary>
        /// <param name="date">A <see cref="DateTime"/> value.</param>
        /// <param name="timeZone">A <see cref="TimeZoneInfo"/> value.</param>
        /// <param name="precision">Precision in terms of significant digits. If zero
        /// then set to maximum (<see cref="PrecisionMax"/>).</param>
        /// <remarks>
        /// <para>The TimeZoneTag component is derived using <see cref="TimeZoneInfo.GetUtcOffset(DateTime)"/>
        /// </para>
        /// <para>If precision is zero, the precision is detected by the number of trailing zeros
        /// after the seconds decimal point. The lowest precision detected is <see cref="PrecisionSecond"/>.
        /// See <see cref="DetectPrecision(DateTime)"/>.
        /// </para>
        /// </remarks>
        public DateTag(DateTime date, TimeZoneInfo timeZone, int precision = 0)
            : this(date, new TimeZoneTag(timeZone.GetUtcOffset(date)), precision)
        {
        }

        /// <summary>
        /// Constructs a DateTag from a DateTimeOffset and precision
        /// </summary>
        /// <param name="date">A <see cref="DateTimeOffset"/> value.</param>
        /// <param name="precision">Precision in terms of significant digits. If zero
        /// then set to maximum (<see cref="PrecisionMax"/>).</param>
        /// <remarks>
        /// <para>The TimeZoneTag component is trawn from <see cref="DateTimeOffset.Offset"/>
        /// </para>
        /// <para>If precision is zero, the precision is detected by the number of trailing zeros
        /// after the seconds decimal point. The lowest precision detected is <see cref="PrecisionSecond"/>.
        /// See <see cref="DetectPrecision(DateTime)"/>.
        /// </para>
        /// </remarks>
        public DateTag(DateTimeOffset date, int precision = 0)
            : this(new DateTime(date.Ticks, DateTimeKind.Local), new TimeZoneTag(date.Offset), precision)
        {
        }

        /// <summary>
        /// If TimeZone.Kind is <see cref="TimeZoneKind.ForceLocal"/> or <see cref="TimeZoneKind.ForceUtc"/>
        /// resolves the timezone offset according to the default passed in. Else returns the DateTag
        /// unchanged.
        /// </summary>
        /// <param name="defaultTimeZone">The default <see cref="TimeZoneInfo"/> with which to resolve
        /// the timezone if none is already available. Use <see cref="TimeZoneInfo.Local"/> for the
        /// current system timezone.
        /// </param>
        /// <returns>A <see cref="TimeZoneTag"/> in which the TimeZone offset has been resolved.</returns>
        /// <remarks>
        /// <para>The time zone portion of a WTCDTF date-time string may be "Z" in which case
        /// <see cref="TimeZoneTag.Kind"/> will be set to <see cref="TimeZoneTag.ForceUtc"/>. The time zone
        /// portion may be blank in which case <see cref="TimeZoneTag.Kind"/> will be set to
        /// <see cref="TimeZoneTag.ForceLocal"/>. Under either of these conditions it may be important
        /// to resolve the timezone before performing date/time operations.
        /// </para>
        /// <para>This method updates the timezone to the default ONLY if the existing value
        /// is either <see cref="TimeZoneKind.ForceLocal"/> or <see cref="TimeZoneKind.ForceUtc"/>.
        /// </para>
        /// <para>When <see cref="TimeZone"/> is <see cref="TimeZoneKind.ForceLocal"/>, the value of
        /// <see cref="Date"/> will be the same in the resulting output while the value of <see cref="DateUtc"/>
        /// will be adjusted according to the appropriate timezone offset. When <see cref="TimeZone"/> is
        /// <see cref="TimeZoneKind.ForceUtc"/> then the value of <see cref="DateUtc"/> will be the same
        /// in original and returned values while the value of <see cref="Date"/> will be adjusted according
        /// to the appropriate TimeZone offset.
        /// </para>
        /// </remarks>
        public DateTag ResolveTimeZone(TimeZoneInfo defaultTimeZone)
        {
            if (TimeZone.Kind == TimeZoneKind.Normal) return this;

            if (TimeZone.Kind == TimeZoneKind.ForceUtc)
            {
                return new DateTag(DateUtc,
                    new TimeZoneTag(defaultTimeZone.GetUtcOffset(DateUtc), TimeZoneKind.Normal),
                    Precision);
            }
            else
            {
                System.Diagnostics.Debug.Assert(TimeZone.Kind == TimeZoneKind.ForceLocal);
                return new DateTag(Date,
                    new TimeZoneTag(defaultTimeZone.GetUtcOffset(Date), TimeZoneKind.Normal),
                    Precision);
            }
        }

        /// <summary>
        /// Convert to <see cref="DateTimeOffset"/>
        /// </summary>
        /// <returns>A <see cref="DateTimeOffset"/> matching the local time and timezone offset of the <see cref="DateTag"/>.</returns>
        public DateTimeOffset ToDateTimeOffset()
        {
            return TimeZone.ToDateTimeOffset(Date);
        }

        /// <summary>
        /// Renders a human-friendly string much like "Sat, 8 Dec 2018, 4:25 PM"
        /// </summary>
        /// <param name="format">A custom format string or null to use <see cref="DateTimeFormatInfo.FullDateTimePattern"/>.</param>
        /// <param name="cultureInfo">The CultureInfo for localization purposes or null to use <see cref="CultureInfo.CurrentCulture"/>.</param>
        /// <returns>The human-friendly string</returns>
        /// <remarks>
        /// <para>Local time will be used unless utcDefault is true, or the standard date and
        /// time format string "r", "R", "u" or "U" is specified.
        /// </para>
        /// <para>The result is sensitive to precision, and localized to <paramref name="cultureInfo"/>.
        /// </para>
        /// <para>
        /// </para>
        /// <para>If the timezone is unresolved (<see cref="TimeZone"/> is <see cref="TimeZoneKind.ForceLocal"/> or
        /// <see cref="TimeZoneKind.ForceUtc"/>) then the current system timezone will be used for conversion
        /// when necessary. If that is not the desired behavior then call <see cref="ResolveTimeZone(TimeZoneInfo)"/>
        /// before using this method.
        /// </para>
        /// <para>Use <code>ToString(null)</code> to get default human-friendly formatting. Calling
        /// <code>ToString()</code> will call the default function which returns the metadata
        /// format.
        /// </para>
        /// </remarks>
        public string ToString(string format, CultureInfo cultureInfo = null, bool utcDefault = false)
        {
            bool useUniversalTime = utcDefault;

            if (cultureInfo == null) cultureInfo = CultureInfo.CurrentCulture;
            if (format == null)
            {
                format = string.Concat(cultureInfo.DateTimeFormat.LongDatePattern, " ",
                    cultureInfo.DateTimeFormat.LongTimePattern);
            }
            else if (format.Length == 1)
            {
                // Resolve standard format string into a custom format string
                switch (format[0])
                {
                    case 'd':
                        format = cultureInfo.DateTimeFormat.ShortDatePattern;
                        break;
                    case 'D':
                        format = cultureInfo.DateTimeFormat.LongDatePattern;
                        break;
                    case 'f': // Full date/time pattern (short time)
                        format = string.Concat(cultureInfo.DateTimeFormat.LongDatePattern, " ",
                            cultureInfo.DateTimeFormat.ShortTimePattern);
                        break;
                    case 'F': // Full date/time pattern (long time)
                        format = string.Concat(cultureInfo.DateTimeFormat.LongDatePattern, " ",
                            cultureInfo.DateTimeFormat.LongTimePattern);
                        break;
                    case 'g':
                        format = string.Concat(cultureInfo.DateTimeFormat.ShortDatePattern, " ",
                            cultureInfo.DateTimeFormat.ShortTimePattern);
                        break;
                    case 'G':
                        format = string.Concat(cultureInfo.DateTimeFormat.ShortDatePattern, " ",
                            cultureInfo.DateTimeFormat.LongTimePattern);
                        break;
                    case 'M':
                    case 'm':
                        format = cultureInfo.DateTimeFormat.MonthDayPattern;
                        break;
                    case 'O':
                    case 'o':
                        format = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffffff";
                        break;
                    case 'R':
                    case 'r':
                        useUniversalTime = true;
                        format = cultureInfo.DateTimeFormat.RFC1123Pattern;
                        break;
                    case 's':
                        format = cultureInfo.DateTimeFormat.SortableDateTimePattern;
                        break;
                    case 't':
                        format = cultureInfo.DateTimeFormat.ShortTimePattern;
                        break;
                    case 'T':
                        format = cultureInfo.DateTimeFormat.LongTimePattern;
                        break;
                    case 'u':
                        useUniversalTime = true;
                        format = cultureInfo.DateTimeFormat.SortableDateTimePattern;
                        break;
                    case 'U':
                        useUniversalTime = true;
                        format = string.Concat(cultureInfo.DateTimeFormat.LongDatePattern, " ",
                            cultureInfo.DateTimeFormat.LongTimePattern);
                        break;
                    case 'Y':
                    case 'y':
                        format = cultureInfo.DateTimeFormat.YearMonthPattern;
                        break;
                    default:
                        throw new FormatException("Unknown standard DateTime format string: " + format);
                }
            }

            format = PrecisionAdaptFormatString(Precision, format);

            if (useUniversalTime)
            {
                return ResolveTimeZone(TimeZoneInfo.Local).DateUtc.ToString(format, cultureInfo);
            }
            else
            {
                return ResolveTimeZone(TimeZoneInfo.Local).ToDateTimeOffset().ToString(format, cultureInfo);
            }
        }

        #endregion Constructor and Methods

        #region Standard Methods

        /// <summary>
        /// Formats a <see cref="DateTag"/> into a date metadata value
        /// according to the <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF</see> standard.
        /// </summary>
        /// <remarks>
        /// <para>The W2CDTF format permits partial date-time values. For example "2018" is just a year with no
        /// other information. The <paramref name="precision"/> value indicates how much detail is included
        /// as follows: 4 = year, 6 = month, 8 = day, 10 = hour, 12 = minute, 14 = second, 17 = millisecond, 20 = microsecond,
        /// 21 = tick (100 nanoseconds).</para>
        /// </remarks>
        public override string ToString()
        {
            var date = Date;
            var sb = new StringBuilder();
            sb.AppendFormat("{0:D4}", date.Year);
            if (Precision >= 6)
            {
                sb.AppendFormat("-{0:D2}", date.Month);
            }
            if (Precision >= 8)
            {
                sb.AppendFormat("-{0:D2}", date.Day);
            }
            if (Precision >= 10)
            {
                sb.AppendFormat("T{0:D2}", date.Hour);
            }
            if (Precision >= 12)
            {
                sb.AppendFormat(":{0:D2}", date.Minute);
            }
            if (Precision >= 14)
            {
                sb.AppendFormat(":{0:D2}", date.Second);
            }
            if (Precision > 14)
            {
                int decimals = Precision - 14;
                if (decimals > 7) decimals = 7;
                sb.Append('.');
                long ticks = date.Ticks % c_ticksPerSecond;
                long pow = c_ticksPerSecond / 10;
                for (int i = 0; i < decimals; ++i)
                {
                    sb.Append((char)('0' + (ticks / pow) % 10));
                    pow /= 10;
                }
            }
            if (TimeZone.Kind == TimeZoneKind.Normal || TimeZone.Kind == TimeZoneKind.ForceUtc)
            {
                sb.Append(TimeZone.ToString());
            }

            return sb.ToString();
        }

        public bool Equals(DateTag obj)
        {
            if (obj == null) return false;
            return m_dateTicks == obj.m_dateTicks && TimeZone.Equals(obj.TimeZone) && Precision == obj.Precision;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as DateTag);
        }

        public override int GetHashCode()
        {
            return Date.GetHashCode()
                ^ TimeZone.GetHashCode()
                ^ Precision.GetHashCode();
        }

        #endregion
    }
}
