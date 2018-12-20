using System;
using System.Collections.Generic;
using System.Text;
using FileMeta;

namespace FMPhotoFinish
{
    static class TimeZoneParser
    {
        static KeyValuePair<string, string>[] s_tzAbbreviations = new KeyValuePair<string, string>[]
        {
            new KeyValuePair<string, string>("Hawaii", "Hawaiian Standard Time"),
            new KeyValuePair<string, string>("Alaska", "Alaskan Standard Time"),
            new KeyValuePair<string, string>("PT", "Pacific Standard Time"),
            new KeyValuePair<string, string>("MT", "Mountain Standard Time"),
            new KeyValuePair<string, string>("Arizona", "US Mountain Standard Time"),
            new KeyValuePair<string, string>("CT", "Central Standard Time"),
            new KeyValuePair<string, string>("ET", "Eastern Standard Time"),
            new KeyValuePair<string, string>("UTC", "UTC"),
            new KeyValuePair<string, string>("GMT", "GMT Standard Time")
        };

        static string DisplayName(TimeZoneInfo tz)
        {
            string displayName = tz.DisplayName;
            if (tz.SupportsDaylightSavingTime)
            {
                displayName += $" (DST: {new FileMeta.TimeZoneTag((int)(tz.BaseUtcOffset.TotalMinutes + tz.GetAdjustmentRules()[0].DaylightDelta.TotalMinutes), FileMeta.TimeZoneKind.Normal).ToString()})";
            }
            return displayName;
        }

        public static void ListTimezoneIds()
        {
            Console.WriteLine("Timezones:");
            foreach (var tz in TimeZoneInfo.GetSystemTimeZones())
            {
                Console.WriteLine($"   {string.Concat("\"", tz.Id, "\""),-34} {DisplayName(tz)}");
            }

            Console.WriteLine();

            Console.WriteLine("Timezone Abbreviations (accepted by this tool):");
            foreach (var pair in s_tzAbbreviations)
            {
                try
                {
                    var tz = TimeZoneInfo.FindSystemTimeZoneById(pair.Value);
                    Console.WriteLine($"   {pair.Key,-34} {DisplayName(tz)}");
                }
                catch
                {
                    // Do nothing. FindTimeZoneById throws an exception if the value is not found.
                    // In that case we just skip the abbreviation.
                }
            }

            Console.WriteLine();
            Console.WriteLine("Examples of Explicit Timezone Values (not sensitive to daylight savings):");
            Console.WriteLine("   \"-08:00\"\r\n   \"-05:00\"\r\n   \"+00:00\"\r\n   \"+01:00\"\r\n   \"+10:00\"\r\n");
        }

        public static TimeZoneInfo ParseTimeZoneId(string id)
        {
            // See if explicit timezone offset
            if (id[0] == '-' || id[0] == '+')
            {
                TimeZoneTag tzt;
                if (TimeZoneTag.TryParse(id, out tzt))
                {
                    string idstr = $"(UTC{id})";
                    string name = $"{idstr} Custom";
                    return TimeZoneInfo.CreateCustomTimeZone(idstr, tzt.UtcOffset, name, name);
                }
            }

            // Check the abbreviations
            foreach(var pair in s_tzAbbreviations)
            {
                if (string.Equals(id, pair.Key, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var tz = TimeZoneInfo.FindSystemTimeZoneById(pair.Value);
                        if (tz != null)
                        {
                            return tz;
                        }
                    }
                    catch
                    {
                        // Do nothing. FindTimeZoneById throws an exception if the value is not found (bad design).
                        // In that case we just skip the abbreviation.
                    }
                    break;
                }
            }

            // Look it up directly
            try
            {
                var tz = TimeZoneInfo.FindSystemTimeZoneById(id);
                if (tz != null)
                {
                    return tz;
                }
            }
            catch
            {
                // Do nothing. FindTimeZoneById throws an exception if the value is not found.
            }
            return null;
        }

    }
}
