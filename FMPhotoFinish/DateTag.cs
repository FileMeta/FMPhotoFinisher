﻿/*
---
name: DateTag.cs
description: CodeBit class that represents a Date metadata tag including DateTime, TimeZone, and Precision components. Includes parsing and formatting methods.
url: https://raw.githubusercontent.com/FileMeta/DateTag/master/DateTag.cs
version: 1.0
keywords: CodeBit
dateModified: 2019-01-23
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
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        #endregion

        #region Public Properties

        /// <summary>
        /// The date value. Always in local time unless TimeZoneTag is ForceUtc.
        /// </summary>
        public DateTime Date { get; private set; }

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
        /// if the <paramref name="date"/> <see cref="DateTime.Kind"/> is <see cref="DateTimeKind.Local"/>,
        /// to <see cref="TimeZoneTag.ForceUtc"/> if <see cref="DateTime.Kind"/> is <see cref="DateTimeKind.Utc"/>,
        /// and to <see cref="TimeZoneTag.Unknown"/> if <see cref="DateTime.Kind"/> is <see cref="DateTimeKind.Unspecified"/>.
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

            // Change date to a compatible timezone (does nothing if already compatible).
            else
            {
                switch (timeZone.Kind)
                {
                    case TimeZoneKind.Normal:
                    case TimeZoneKind.ForceLocal:
                        date = timeZone.ToLocal(date);
                        break;

                    case TimeZoneKind.ForceUtc:
                        date = timeZone.ToUtc(date);
                        break;
                }
            }

            // Limit precision to compatible range
            if (precision > PrecisionMax) precision = PrecisionMax;
            if (precision < PrecisionMin) precision = DetectPrecision(date);          

            Date = date;
            TimeZone = timeZone;
            Precision = precision;
        }

        /// <summary>
        /// If TimeZone.Kind is <see cref="TimeZoneKind.ForceLocal"/> or <see cref="TimeZoneKind.ForceUtc"/>
        /// resolves the timezone offset according to the default passed in. Else does nothing.
        /// </summary>
        /// <param name="defaultTimeZone">The default <see cref="TimeZoneInfo"/> with which to resolve
        /// thei timezone. Use <see cref="TimeZoneInfo.Local"/> for the current system timezone.
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
        /// </remarks>
        public DateTag ResolveTimeZone(TimeZoneInfo defaultTimeZone)
        {
            if (TimeZone.Kind == TimeZoneKind.Normal) return this;

            var tz = new TimeZoneTag(defaultTimeZone.GetUtcOffset(Date), TimeZoneKind.Normal);
            return new DateTag(Date, tz, Precision);
        }

        /// <summary>
        /// Convert the <see cref="Date"/> value to UTC (if it is not already).
        /// </summary>
        /// <returns>A <see cref="DateTime"/> with a <see cref="DateTimeKind"/> of <see cref="DateTimeKind.Utc"/>.</returns>
        /// <remarks>
        /// <para>If the <see cref="Date"/> is <see cref="DateTimeKind.Local"/> then the value is converted to
        /// UTC using the offset in <see cref="TimeZone"/>.
        /// </para>
        /// <para>If <see cref="TimeZone.Kind"/> is <see cref="TimeZoneKind.ForceLocal"/> then the returned value
        /// will have unchanged time but with <see cref="DateTime.Kind"/> changed to <see cref="DateTimeKind.Utc"/>.
        /// This may or may not be the needed behavior. A prior call to <see cref="ResolveTimeZone"/> may be necessary
        /// to achieve the desired results.
        /// </para>
        /// </remarks>
        public DateTime ToUtc()
        {
            return TimeZone.ToUtc(Date);
        }

        /// <summary>
        /// Convert the <see cref="Date"/> value to Local (if it is not already).
        /// </summary>
        /// <returns>A <see cref="DateTime"/> with a <see cref="DateTimeKind"/> of <see cref="DateTimeKind.Local"/>.</returns>
        /// <remarks>
        /// <para>If the <see cref="Date"/> is <see cref="DateTimeKind.Utc"/> then the value is converted to
        /// Local using the offset in <see cref="TimeZone"/>. If the Date is already local then the result is unchanged.
        /// </para>
        /// <para>If <see cref="TimeZone.Kind"/> is <see cref="TimeZoneKind.ForceUtc"/> then the returned value
        /// will have unchanged time but with <see cref="DateTime.Kind"/> changed to <see cref="DateTimeKind.Utc"/>.
        /// This may or may not be the needed behavior. A prior call to <see cref="ResolveTimeZone"/> may be necessary
        /// to achieve the desired results.
        /// </para>
        /// </remarks>
        public DateTime ToLocal()
        {
            return TimeZone.ToLocal(Date);
        }

        #endregion Constructor and Methods

        #region Standard Methods

        /// <summary>
        /// Formats a <see cref="DateTag"/> into adate metadata value
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
            if (TimeZone.Kind == TimeZoneKind.Normal && Date.Kind == DateTimeKind.Utc)
            {
                Date = TimeZone.ToLocal(Date);
            }

            var sb = new StringBuilder();
            sb.AppendFormat("{0:D4}", Date.Year);
            if (Precision >= 6)
            {
                sb.AppendFormat("-{0:D2}", Date.Month);
            }
            if (Precision >= 8)
            {
                sb.AppendFormat("-{0:D2}", Date.Day);
            }
            if (Precision >= 10)
            {
                sb.AppendFormat("T{0:D2}", Date.Hour);
            }
            if (Precision >= 12)
            {
                sb.AppendFormat(":{0:D2}", Date.Minute);
            }
            if (Precision >= 14)
            {
                sb.AppendFormat(":{0:D2}", Date.Second);
            }
            if (Precision > 14)
            {
                int decimals = Precision - 14;
                if (decimals > 7) decimals = 7;
                sb.Append('.');
                long ticks = Date.Ticks % c_ticksPerSecond;
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
            return Date.Equals(obj.Date) && TimeZone.Equals(obj.TimeZone) && Precision.Equals(obj.Precision);
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
