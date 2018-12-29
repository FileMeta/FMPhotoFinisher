/*
This will eventually be a CodeBit. The class manages Date metadata field values.
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

        public const int MinPrecision = 4;
        public const int MaxPrecision = 21;
        public static readonly DateTime ZeroDate = new DateTime(0L);

        const long c_ticksPerSecond = 10000000;

        #endregion Constants

        #region Static Methods

        /// <summary>
        /// Parses a metadata date tag into <see cref="DateTime"/>, <see cref="TimeZoneTag"/>, and significant digits.
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
            result = new DateTag(ZeroDate, TimeZoneTag.Unknown, 0);

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

            if (!int.TryParse(dateTag.Substring(0, 4), out year)
                || year < 1 || year > 9999) return false;
            int precision = 4; // year
            pos = 4;
            if (dateTag.Length > 5 && dateTag[4] == '-')
            {
                if (!int.TryParse(dateTag.Substring(5, 2), out month)
                    || month < 1 || month > 12) return false;
                precision = 6; // month
                pos = 7;
                if (dateTag.Length > 8 && dateTag[7] == '-')
                {
                    if (!int.TryParse(dateTag.Substring(8, 2), out day)
                        || day < 1 || day > 31) return false;
                    precision = 8; // day
                    pos = 10;
                    if (dateTag.Length > 11 && (dateTag[10] == 'T' || dateTag[10] == ' ')) // Even though W3CDTF and ISO 8601 specify 'T' separating date and time, tolerate a space as an alternative.
                    {
                        if (!int.TryParse(dateTag.Substring(11, 2), out hour)
                            || hour < 0 || hour > 23) return false;
                        precision = 10; // hour
                        pos = 13;
                        if (dateTag.Length > 14 && dateTag[13] == ':')
                        {
                            if (!int.TryParse(dateTag.Substring(14, 2), out minute)
                                || minute < 0 || minute > 59) return false;
                            precision = 12; // minute
                            pos = 16;
                            if (dateTag.Length > 17 && dateTag[16] == ':')
                            {
                                if (!int.TryParse(dateTag.Substring(17, 2), out second)
                                    || second < 0 || second > 59) return false;
                                precision = 14; // second
                                pos = 19;
                                if (dateTag.Length > 20 && dateTag[19] == '.')
                                {
                                    ++pos;
                                    int anchor = pos;
                                    while (pos < dateTag.Length && char.IsDigit(dateTag[pos]))
                                        ++pos;

                                    precision = 14 + (pos - anchor);
                                    if (precision > MaxPrecision)
                                        precision = MaxPrecision;

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

        public DateTag(DateTime date, TimeZoneTag timeZone, int precision)
        {
            Date = date;
            TimeZone = timeZone;

            if (precision > MaxPrecision) precision = MaxPrecision;
            if (precision < MinPrecision) precision = MinPrecision;
            Precision = precision;
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

        #region Test

#if DEBUG

        /* When this is converted into a CodeBit, the tests should be
         * automated (verifying the result rather than visual verification)
         * and included as the unit test for the dedicated project. */

        public static void PerformTest()
        {
            TestTagParse("2018");
            TestTagParse("2018-12");
            TestTagParse("2018-12-28");
            TestTagParse("2018-12-28T10");
            TestTagParse("2018-12-28T10-05:00");
            TestTagParse("2018-12-28T10+08:00");
            TestTagParse("2018-12-28T10:59");
            TestTagParse("2018-12-28T10:59-05");
            TestTagParse("2018-12-28T10:59+08");
            TestTagParse("2018-12-28T10:59:45");
            TestTagParse("2018-12-28T10:59:45-5");
            TestTagParse("2018-12-28T10:59:45+8");
            TestTagParse("2018-12-28T10:59:45.123");
            TestTagParse("2018-12-28T10:59:45.123-5");
            TestTagParse("2018-12-28T10:59:45.123+8");
            TestTagParse("2018-12-28T10:59:45.1");
            TestTagParse("2018-12-28T10:59:45.12");
            TestTagParse("2018-12-28T10:59:45.123");
            TestTagParse("2018-12-28T10:59:45.1234");
            TestTagParse("2018-12-28T10:59:45.12345");
            TestTagParse("2018-12-28T10:59:45.123456");
            TestTagParse("2018-12-28T10:59:45.1234567");
            TestTagParse("2018-12-28T10:59:45.12345678");
            TestTagParse("2018-12-28T10:59:45.12345678+08:23");
        }

        static void TestTagParse(string s)
        {
            Console.WriteLine(s);
            DateTag date;
            if (!TryParse(s, out date))
            {
                Console.WriteLine("Fail");
                return;
            }
            Console.WriteLine($"   Date: {date.Date.ToString("o")}");
            Console.WriteLine($"   Fmt:  {date}");
            Console.WriteLine($"   TZ:   {date.TimeZone}");
            Console.WriteLine($"   Pre:  {date.Precision}");
        }

#endif

        #endregion Tests

    }
}
