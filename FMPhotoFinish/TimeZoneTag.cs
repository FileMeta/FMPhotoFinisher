/*
This will eventually be a CodeBit. The class manages TimeZone metadata field values.
*/
using System;
using System.Globalization;
using System.Text;

namespace FileMeta
{
    enum TimeZoneKind : int
    {
        /// <summary>
        /// Normal timezone - specifies offset from UTC.
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Date-Time field should be treated as local time regardless of whether its
        /// definition indicates UTC or local.
        /// </summary>
        ForceLocal = 1,

        /// <summary>
        /// Date-Tiem field should be treated as UTC time regardless of whether its
        /// definition indicates UTC or local.
        /// </summary>
        ForceUtc = 2
    }

    /// <summary>
    /// TimeZoneMetadata represents the timezone portion of a Date (date-time) metadata
    /// tag or of a dedicated timezone metadata tag.
    /// </summary>
    /// <remarks>
    /// <para>An ideal date field follows the <see cref="https://www.w3.org/TR/NOTE-datetime">W3CDTF profile
    /// of ISO 8601</see> and includes the timezone informaiton, Here is an example:
    /// </para>
    /// <code>2018-11-28T13:25:04-05:00</code>
    /// <para>In this case, the value indicates 28 November 2018 at 1:25:04 pm in the Eastern Standard time zone
    /// (UTC - 5 hours). The TimeZoneMetadata class represents the timezone portion of such a date-time.
    /// The <see cref="TryParseDate"/> method will parse a date field in this format and return
    /// <see cref="System.DateTime"/> and <see cref="TimeZoneTag"/> values.
    /// </para>
    /// <para>For the example above, <see cref="Kind"/> would be <see cref="TimeZoneKind.Normal"/> and
    /// <see cref="UtcOffset"/> would be negative five hours.
    /// </para>
    /// <para>The recommended format for Date includes the timezone but it is frequently omitted. When no
    /// time zone is present then <see cref="Kind"/> is <see cref="TimeZoneKind.Unknown"/>. W3CDTF also allows
    /// a time zone suffix of "Z" which means the time is UTC. In that case, <see cref="Kind"/> is
    /// <see cref="TimeZoneKind.ForceUtc"/>.
    /// </para>
    /// <para>Most existing metadata formats do not support explicit timezone information. Often the date-time
    /// is stored as a binary value or a formatted value without timezone. For those formats, a separate
    /// "timezone" tag may be used to augment the existing field with timezone information. Existing date
    /// values are usually defined (in the format documentation) as either local time or UTC. Regardless of
    /// the field definition, the timezone tag should indicate the difference between local time and UTC. So,
    /// for example, "-05:00" means local time is UTC minus five hours. Minutes are included because some
    /// timezones are offset by a half hour. The sign SHOULD always be included (e.g. "+08:00" for UTC plus
    /// eight hours).
    /// </para>
    /// <para>There are two special values for the "timezone" metadata field:
    /// </para>
    /// <para>"0" indicates that the timezone is unknown and that all fields should be treated as local time
    /// regardless of the documented field definition. This is represented by <see cref="Kind"/> being set
    /// to <see cref="TimeZoneKind.ForceLocal"/>. This is common for cameras that produce video in .mov
    /// or .mp4 format. The "date_created" metadata field for those file formats is defined as being UTC.
    /// But such cameras often do not have a timezone setting and, consequently, they store the local time
    /// in the "date_created" field.
    /// </para>
    /// <para>"Z" indicates that the timezone is unknown and that all fields should be treated as UTC
    /// regardless of the documented field definition. This is represented by <see cref="Kind"/> being set
    /// to <see cref="TimeZoneKind.ForceUtc"/>.
    /// </para>
    /// <para>If <see cref="Kind"/> is other than <see cref="TimeZoneKind.Normal"/> then
    /// <see cref="UtcOffset"/> must be zero.
    /// </para>
    /// <para>TimeZoneTag is immutable.
    /// </para>
    /// </remarks>
    class TimeZoneTag
    {
        #region Constants

        const string c_local = "0";
        const string c_utc = "Z";
        const long c_ticksPerSecond = 10000000;
        const long c_ticksPerMinute = 60 * c_ticksPerSecond;

        #endregion Constants

        #region Static Methods and Properties

        /// <summary>
        /// Parses a timezone string into a TimeZoneTag instance.
        /// </summary>
        /// <param name="s">The timezone string to parse.</param>
        /// <param name="result">The parsed timezone.</param>
        /// <returns>True if successful, else false.</returns>
        /// <remarks>
        /// <para>See <see cref="TimeZoneTag"/> for details about valid values.
        /// </para>
        /// <para>Example timezone values:</para>
        /// <para>  "-05:00" (UTC minus 5 hours)</para>
        /// <para>  "+06:00" (UTC plus 6 hours)</para>
        /// <para>  "+09:30" (UTC plus 9 1/2 hours)</para>
        /// <para>  "Z"      (UTC. Offset to local is unknown.)</para>
        /// <para>  "0"      (Local. Offset to UTC is unknwon.)</para>
        /// <para>Tolerable timezone values:</para>
        /// <para>  "-5"     (UTC minus 5 hours)</para>
        /// <para>  "+6      (UTC plus 6 hours)</para>
        /// </remarks>
        public static bool TryParse(string timezoneTag, out TimeZoneTag result)
        {
            if (string.IsNullOrEmpty(timezoneTag))
            {
                result = Zero;
                return false;
            }
            if (timezoneTag.Equals(c_local, StringComparison.Ordinal))
            {
                result = ForceLocal;
                return true;
            }
            if (timezoneTag.Equals(c_utc, StringComparison.Ordinal))
            {
                result = ForceUtc;
                return true;
            }

            result = Zero;
            if (timezoneTag.Length < 2) return false;

            bool negative;
            if (timezoneTag[0] == '+')
            {
                negative = false;
            }
            else if (timezoneTag[0] == '-')
            {
                negative = true;
            }
            else
            {
                return false;
            }

            var parts = timezoneTag.Substring(1).Split(':');
            if (parts.Length < 1 || parts.Length > 2) return false;
            int hours;
            if (!int.TryParse(parts[0], out hours)) return false;
            if (hours < 0 || hours > 23) return false;

            int minutes = 0;
            if (parts.Length > 1)
            {
                if (!int.TryParse(parts[1], out minutes)) return false;
                if (minutes < 0 || minutes > 59) return false;
            }

            int totalMinutes = hours * 60 + minutes;
            if (negative) totalMinutes = -totalMinutes;
            result = new TimeZoneTag(totalMinutes, TimeZoneKind.Normal);
            return true;
        }

        #endregion Static Methods

        public static readonly TimeZoneTag Zero = new TimeZoneTag(0, TimeZoneKind.Normal);
        public static readonly TimeZoneTag ForceLocal = new TimeZoneTag(0, TimeZoneKind.ForceLocal);
        public static readonly TimeZoneTag ForceUtc = new TimeZoneTag(0, TimeZoneKind.ForceUtc);

        private int m_offset;   // Timezone offset in minutes

        #region Properties

        /// <summary>
        /// The <see cref="TimeZoneKind"/> of this TimeZoneTag
        /// </summary>
        public TimeZoneKind Kind { get; private set; }

        /// <summary>
        /// Offset from UTC as a <see cref="TimeSpan"/>.
        /// </summary>
        /// <remarks>
        /// Add this value to a UTC time in order to get a local time. Likewise,
        /// substract this value from a local time to get a UTC time. However, it
        /// is preferable to use the <see cref="ToLocal"/> and <see cref="ToUtc"/>
        /// methods because they are sensitive to the <see cref="DateTime.Kind"/>
        /// value on the inbound value and correctly set the DateTime.Kind value
        /// on the result.</remarks>
        public TimeSpan UtcOffset { get { return new TimeSpan(UtcOffsetTicks); } }

        /// <summary>
        /// Offset from UTC in minutes.
        /// </summary>
        /// <seealso cref="UtcOffset"/>
        public int UtcOffsetMinutes {  get { return m_offset; } }

        /// <summary>
        /// Offset from UTC in ticks.
        /// </summary>
        /// <seealso cref="UtcOffset"/>
        public long UtcOffsetTicks { get { return m_offset * c_ticksPerMinute; } }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Constructs a TimeZoneTag.
        /// </summary>
        /// <param name="offsetMinutes">Offset from UTC in minutes.</param>
        /// <param name="kind">The <see cref="TimeZoneKind"/>.</param>
        /// <remarks>If <paramref name="kind"/> is other than <see cref="TimeZoneKind.Normal"/>
        /// then <paramref name="offsetMinutes"/> is forced to zero.
        /// </remarks>
        public TimeZoneTag(int offsetMinutes, TimeZoneKind kind)
        {
            m_offset = (kind == TimeZoneKind.Normal) ? offsetMinutes : 0;
            Kind = kind;
        }

        /// <summary>
        /// Constructs a TimeZoneTag.
        /// </summary>
        /// <param name="offsetTicks">Offset from UTC in ticks.</param>
        /// <param name="kind">The <see cref="TimeZoneKind"/>.</param>
        /// <remarks>If <paramref name="kind"/> is other than <see cref="TimeZoneKind.Normal"/>
        /// then <paramref name="offsetTicks"/> is forced to zero.
        /// </remarks>
        public TimeZoneTag(long offsetTicks, TimeZoneKind kind)
        {
            m_offset = (kind == TimeZoneKind.Normal) ? (int)(offsetTicks / c_ticksPerMinute) : 0;
            Kind = kind;
        }

        /// <summary>
        /// Constructs a TimeZoneTag.
        /// </summary>
        /// <param name="offset">Offset from UTC.</param>
        /// <param name="kind">The <see cref="TimeZoneKind"/>.</param>
        /// <remarks>If <paramref name="kind"/> is <see cref="TimeZoneKind.ForceUtc"/>
        /// or <see cref="TimeZoneKind.ForceLocal"/> then <paramref name="offsetMinutes"/>
        /// must be zero. Other values will throw an exception.</remarks>
        /// <remarks>If <paramref name="kind"/> is other than <see cref="TimeZoneKind.Normal"/>
        /// then <paramref name="offset"/> is forced to zero.
        /// </remarks>
        public TimeZoneTag(TimeSpan offset, TimeZoneKind kind)
        {
            m_offset = (kind == TimeZoneKind.Normal) ? (int)(offset.Ticks / c_ticksPerMinute) : 0;
            Kind = kind;
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// Convert a <see cref="DateTime"/> to local time if it is not already.
        /// </summary>
        /// <param name="date">The value to convert.</param>
        /// <returns>A <see cref="DateTime"/> in the local timezone with <see cref="DateTime.Kind"/>
        /// set to <see cref="DateTimeKind.Local"/>.</returns>
        /// <remarks>
        /// <para>If <see cref="DateTime.Kind"/> on the inbound value is set to <see cref="DateTimeKind.Local"/>
        /// then this method returns the inbound value unchanged.
        /// </para>
        /// <para>If <see cref="DateTime.Kind"/> on the inbound value is set to <see cref="DateTimeKind.Utc"/>
        /// or to <see cref="DateTimeKind.Unspecified"/> then the value is converted to local time
        /// by adding the time zone offset and setting <see cref="DateTime.Kind"/> to <see cref="DateTimeKind.Local"/>;
        /// </para>
        /// <para>Note that if <see cref="Kind">TimeZoneInfo.Kind</see> is other than <see cref="TimeZoneKind.Normal"/>
        /// then the offset will be zero.
        /// </para>
        /// </remarks>
        public DateTime ToLocal(DateTime date)
        {
            if (date.Kind == DateTimeKind.Local) return date;
            return new DateTime(date.Ticks + (m_offset * c_ticksPerMinute), DateTimeKind.Local);
        }

        /// <summary>
        /// Convert a <see cref="DateTime"/> to UTC time if it is not already.
        /// </summary>
        /// <param name="date">The value to convert.</param>
        /// <returns>A <see cref="DateTime"/> in UTC with <see cref="DateTime.Kind"/>
        /// set to <see cref="DateTimeKind.Utc"/>.</returns>
        /// <remarks>
        /// <para>If <see cref="DateTime.Kind"/> on the inbound value is set to <see cref="DateTimeKind.Utc"/>
        /// then this method returns the inbound value unchanged.
        /// </para>
        /// <para>If <see cref="DateTime.Kind"/> on the inbound value is set to <see cref="DateTimeKind.Local"/>
        /// or to <see cref="DateTimeKind.Unspecified"/> then the value is converted to UTC
        /// by subtracting the time zone offset and setting <see cref="DateTime.Kind"/> to <see cref="DateTimeKind.Utc"/>;
        /// </para>
        /// <para>Note that if <see cref="Kind">TimeZoneInfo.Kind</see> is other than <see cref="TimeZoneKind.Normal"/>
        /// then the offset will be zero.
        /// </para>
        /// </remarks>
        public DateTime ToUtc(DateTime date)
        {
            if (date.Kind == DateTimeKind.Utc) return date;
            return new DateTime(date.Ticks - (m_offset * c_ticksPerMinute), DateTimeKind.Utc);
        }

        #endregion

        #region Standard Methods

        public override string ToString()
        {
            if (Kind == TimeZoneKind.ForceLocal) return c_local;
            if (Kind == TimeZoneKind.ForceUtc) return c_utc;

            int minutes = m_offset;
            char sign;
            if (minutes < 0)
            {
                sign = '-';
                minutes = -minutes;
            }
            else
            {
                sign = '+';
            }
            return String.Format(CultureInfo.InvariantCulture, "{0}{1:d2}:{2:d2}", sign, minutes / 60, minutes % 60);
        }

        public override int GetHashCode()
        {
            return Kind.GetHashCode() ^ m_offset.GetHashCode();
        }

        public bool Equals(TimeZoneTag other)
        {
            if (other == null) return false;

            return Kind == other.Kind && m_offset == other.m_offset;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as TimeZoneTag);
        }

        #endregion Standard Methods
    }
}
