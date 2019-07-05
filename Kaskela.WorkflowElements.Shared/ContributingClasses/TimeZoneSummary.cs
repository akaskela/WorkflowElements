using System.Collections.Generic;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.ContributingClasses
{
    public class TimeZoneSummary
    {
        public static TimeZoneSummary RetrieveTimeZoneByIndex(int index)
        {
            return RetrieveTimeZones().FirstOrDefault(tzs => tzs.MicrosoftIndex == index);
        }
        public static TimeZoneSummary RetrieveTimeZoneByOptionSetValue(int optionSetValue)
        {
            return RetrieveTimeZones().FirstOrDefault(tzs => tzs.OptionSetValue == optionSetValue);
        }
        public static List<TimeZoneSummary> RetrieveTimeZones()
        {
            return new List<TimeZoneSummary>()
            {
                new TimeZoneSummary { MicrosoftIndex = 0, Id ="Dateline Standard Time", FullName = "(GMT-12:00) International Date Line West", OptionSetValue = 222540002 },
                new TimeZoneSummary { MicrosoftIndex = 1, Id ="Samoa Standard Time", FullName = "(GMT-11:00) Midway Island, Samoa", OptionSetValue = 222540003 },
                new TimeZoneSummary { MicrosoftIndex = 2, Id ="Hawaiian Standard Time", FullName = "(GMT-10:00) Hawaii", OptionSetValue = 222540004 },
                new TimeZoneSummary { MicrosoftIndex = 3, Id ="Alaskan Standard Time", FullName = "(GMT-09:00) Alaska", OptionSetValue = 222540005 },
                new TimeZoneSummary { MicrosoftIndex = 4, Id ="Pacific Standard Time", FullName = "(GMT-08:00) Pacific Time (US and Canada); Tijuana", OptionSetValue = 222540006 },
                new TimeZoneSummary { MicrosoftIndex = 10, Id ="Mountain Standard Time", FullName = "(GMT-07:00) Mountain Time (US and Canada)", OptionSetValue = 222540007 },
                new TimeZoneSummary { MicrosoftIndex = 13, Id ="Mexico Standard Time 2", FullName = "(GMT-07:00) Chihuahua, La Paz, Mazatlan", OptionSetValue = 222540008 },
                new TimeZoneSummary { MicrosoftIndex = 15, Id ="U.S. Mountain Standard Time", FullName = "(GMT-07:00) Arizona", OptionSetValue = 222540009 },
                new TimeZoneSummary { MicrosoftIndex = 20, Id ="Central Standard Time", FullName = "(GMT-06:00) Central Time (US and Canada", OptionSetValue = 222540010 },
                new TimeZoneSummary { MicrosoftIndex = 25, Id ="Canada Central Standard Time", FullName = "(GMT-06:00) Saskatchewan", OptionSetValue = 222540011 },
                new TimeZoneSummary { MicrosoftIndex = 30, Id ="Mexico Standard Time", FullName = "(GMT-06:00) Guadalajara, Mexico City, Monterrey", OptionSetValue = 222540012 },
                new TimeZoneSummary { MicrosoftIndex = 33, Id ="Central America Standard Time", FullName = "(GMT-06:00) Central America", OptionSetValue = 222540013 },
                new TimeZoneSummary { MicrosoftIndex = 35, Id ="Eastern Standard Time", FullName = "(GMT-05:00) Eastern Time (US and Canada)", OptionSetValue = 222540014 },
                new TimeZoneSummary { MicrosoftIndex = 40, Id ="U.S. Eastern Standard Time", FullName = "(GMT-05:00) Indiana (East)", OptionSetValue = 222540015 },
                new TimeZoneSummary { MicrosoftIndex = 45, Id ="S.A. Pacific Standard Time", FullName = "(GMT-05:00) Bogota, Lima, Quito", OptionSetValue = 222540016 },
                new TimeZoneSummary { MicrosoftIndex = 50, Id ="Atlantic Standard Time", FullName = "(GMT-04:00) Atlantic Time (Canada)", OptionSetValue = 222540017 },
                new TimeZoneSummary { MicrosoftIndex = 55, Id ="S.A. Western Standard Time", FullName = "(GMT-04:00) Caracas, La Paz", OptionSetValue = 222540018 },
                new TimeZoneSummary { MicrosoftIndex = 56, Id ="Pacific S.A. Standard Time", FullName = "(GMT-04:00) Santiago", OptionSetValue = 222540019 },
                new TimeZoneSummary { MicrosoftIndex = 60, Id ="Newfoundland and Labrador Standard Time", FullName = "(GMT-03:30) Newfoundland and Labrador", OptionSetValue = 222540020 },
                new TimeZoneSummary { MicrosoftIndex = 65, Id ="E. South America Standard Time", FullName = "(GMT-03:00) Brasilia", OptionSetValue = 222540021 },
                new TimeZoneSummary { MicrosoftIndex = 70, Id ="S.A. Eastern Standard Time", FullName = "(GMT-03:00) Buenos Aires, Georgetown", OptionSetValue = 222540022 },
                new TimeZoneSummary { MicrosoftIndex = 73, Id ="Greenland Standard Time", FullName = "(GMT-03:00) Greenland", OptionSetValue = 222540023 },
                new TimeZoneSummary { MicrosoftIndex = 75, Id ="Mid-Atlantic Standard Time", FullName = "(GMT-02:00) Mid-Atlantic", OptionSetValue = 222540024 },
                new TimeZoneSummary { MicrosoftIndex = 80, Id ="Azores Standard Time", FullName = "(GMT-01:00) Azores", OptionSetValue = 222540025 },
                new TimeZoneSummary { MicrosoftIndex = 83, Id ="Cape Verde Standard Time", FullName = "(GMT-01:00) Cape Verde Islands", OptionSetValue = 222540026 },
                new TimeZoneSummary { MicrosoftIndex = 85, Id ="GMT Standard Time", FullName = "(GMT) Greenwich Mean Time: Dublin, Edinburgh, Lisbon, London", OptionSetValue = 222540027 },
                new TimeZoneSummary { MicrosoftIndex = 90, Id ="Greenwich Standard Time", FullName = "(GMT+00:00) Casablanca", OptionSetValue = 222540028 },
                new TimeZoneSummary { MicrosoftIndex = 95, Id ="Central Europe Standard Time", FullName = "(GMT+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague", OptionSetValue = 222540029 },
                new TimeZoneSummary { MicrosoftIndex = 100, Id ="Central European Standard Time", FullName = "(GMT+01:00) Sarajevo, Skopje, Warsaw, Zagreb", OptionSetValue = 222540030 },
                new TimeZoneSummary { MicrosoftIndex = 105, Id ="Romance Standard Time", FullName = "(GMT+01:00) Brussels, Copenhagen, Madrid, Paris", OptionSetValue = 222540031 },
                new TimeZoneSummary { MicrosoftIndex = 110, Id ="W. Europe Standard Time", FullName = "(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna", OptionSetValue = 222540032 },
                new TimeZoneSummary { MicrosoftIndex = 113, Id ="W. Central Africa Standard Time", FullName = "(GMT+01:00) West Central Africa", OptionSetValue = 222540033 },
                new TimeZoneSummary { MicrosoftIndex = 115, Id ="E. Europe Standard Time", FullName = "(GMT+02:00) Bucharest", OptionSetValue = 222540034 },
                new TimeZoneSummary { MicrosoftIndex = 120, Id ="Egypt Standard Time", FullName = "(GMT+02:00) Cairo", OptionSetValue = 222540035 },
                new TimeZoneSummary { MicrosoftIndex = 125, Id ="FLE Standard Time", FullName = "(GMT+02:00) Helsinki, Kiev, Riga, Sofia, Tallinn, Vilnius", OptionSetValue = 222540036 },
                new TimeZoneSummary { MicrosoftIndex = 130, Id ="GTB Standard Time", FullName = "(GMT+02:00) Athens, Istanbul, Minsk", OptionSetValue = 222540037 },
                new TimeZoneSummary { MicrosoftIndex = 135, Id ="Israel Standard Time", FullName = "(GMT+02:00) Jerusalem", OptionSetValue = 222540038 },
                new TimeZoneSummary { MicrosoftIndex = 140, Id ="South Africa Standard Time", FullName = "(GMT+02:00) Harare, Pretoria", OptionSetValue = 222540039 },
                new TimeZoneSummary { MicrosoftIndex = 145, Id ="Russian Standard Time", FullName = "(GMT+03:00) Moscow, St. Petersburg, Volgograd", OptionSetValue = 222540040 },
                new TimeZoneSummary { MicrosoftIndex = 150, Id ="Arab Standard Time", FullName = "(GMT+03:00) Kuwait, Riyadh", OptionSetValue = 222540041 },
                new TimeZoneSummary { MicrosoftIndex = 155, Id ="E. Africa Standard Time", FullName = "(GMT+03:00) Nairobi", OptionSetValue = 222540042 },
                new TimeZoneSummary { MicrosoftIndex = 158, Id ="Arabic Standard Time", FullName = "(GMT+03:00) Baghdad", OptionSetValue = 222540043 },
                new TimeZoneSummary { MicrosoftIndex = 160, Id ="Iran Standard Time", FullName = "(GMT+03:30) Tehran", OptionSetValue = 222540044 },
                new TimeZoneSummary { MicrosoftIndex = 165, Id ="Arabian Standard Time", FullName = "(GMT+04:00) Abu Dhabi, Muscat", OptionSetValue = 222540045 },
                new TimeZoneSummary { MicrosoftIndex = 170, Id ="Caucasus Standard Time", FullName = "(GMT+04:00) Baku, Tbilisi, Yerevan", OptionSetValue = 222540046 },
                new TimeZoneSummary { MicrosoftIndex = 175, Id ="Transitional Islamic State of Afghanistan Standard Time", FullName = "(GMT+04:30) Kabul", OptionSetValue = 222540047 },
                new TimeZoneSummary { MicrosoftIndex = 180, Id ="Ekaterinburg Standard Time", FullName = "(GMT+05:00) Ekaterinburg", OptionSetValue = 222540048 },
                new TimeZoneSummary { MicrosoftIndex = 185, Id ="West Asia Standard Time", FullName = "(GMT+05:00) Islamabad, Karachi, Tashkent", OptionSetValue = 222540049 },
                new TimeZoneSummary { MicrosoftIndex = 190, Id ="India Standard Time", FullName = "(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi", OptionSetValue = 222540050 },
                new TimeZoneSummary { MicrosoftIndex = 193, Id ="Nepal Standard Time", FullName = "(GMT+05:45) Kathmandu", OptionSetValue = 222540051 },
                new TimeZoneSummary { MicrosoftIndex = 195, Id ="Central Asia Standard Time", FullName = "(GMT+06:00) Astana, Dhaka", OptionSetValue = 222540052 },
                new TimeZoneSummary { MicrosoftIndex = 200, Id ="Sri Lanka Standard Time", FullName = "(GMT+06:00) Sri Jayawardenepura", OptionSetValue = 222540053 },
                new TimeZoneSummary { MicrosoftIndex = 201, Id ="N. Central Asia Standard Time", FullName = "(GMT+06:00) Almaty, Novosibirsk", OptionSetValue = 222540054 },
                new TimeZoneSummary { MicrosoftIndex = 203, Id ="Myanmar Standard Time", FullName = "(GMT+06:30) Yangon Rangoon", OptionSetValue = 222540055 },
                new TimeZoneSummary { MicrosoftIndex = 205, Id ="S.E. Asia Standard Time", FullName = "(GMT+07:00) Bangkok, Hanoi, Jakarta", OptionSetValue = 222540056 },
                new TimeZoneSummary { MicrosoftIndex = 207, Id ="North Asia Standard Time", FullName = "(GMT+07:00) Krasnoyarsk", OptionSetValue = 222540057 },
                new TimeZoneSummary { MicrosoftIndex = 210, Id ="China Standard Time", FullName = "(GMT+08:00) Beijing, Chongqing, Hong Kong SAR, Urumqi", OptionSetValue = 222540058 },
                new TimeZoneSummary { MicrosoftIndex = 215, Id ="Singapore Standard Time", FullName = "(GMT+08:00) Kuala Lumpur, Singapore", OptionSetValue = 222540059 },
                new TimeZoneSummary { MicrosoftIndex = 220, Id ="Taipei Standard Time", FullName = "(GMT+08:00) Taipei", OptionSetValue = 222540060 },
                new TimeZoneSummary { MicrosoftIndex = 225, Id ="W. Australia Standard Time", FullName = "(GMT+08:00) Perth", OptionSetValue = 222540061 },
                new TimeZoneSummary { MicrosoftIndex = 227, Id ="North Asia East Standard Time", FullName = "(GMT+08:00) Irkutsk, Ulaanbaatar", OptionSetValue = 222540062 },
                new TimeZoneSummary { MicrosoftIndex = 230, Id ="Korea Standard Time", FullName = "(GMT+09:00) Seoul", OptionSetValue = 222540063 },
                new TimeZoneSummary { MicrosoftIndex = 235, Id ="Tokyo Standard Time", FullName = "(GMT+09:00) Osaka, Sapporo, Tokyo", OptionSetValue = 222540064 },
                new TimeZoneSummary { MicrosoftIndex = 240, Id ="Yakutsk Standard Time", FullName = "(GMT+09:00) Yakutsk", OptionSetValue = 222540065 },
                new TimeZoneSummary { MicrosoftIndex = 245, Id ="A.U.S. Central Standard Time", FullName = "(GMT+09:30) Darwin", OptionSetValue = 222540066 },
                new TimeZoneSummary { MicrosoftIndex = 250, Id ="Cen. Australia Standard Time", FullName = "(GMT+09:30) Adelaide", OptionSetValue = 222540067 },
                new TimeZoneSummary { MicrosoftIndex = 255, Id ="A.U.S. Eastern Standard Time", FullName = "(GMT+10:00) Canberra, Melbourne, Sydney", OptionSetValue = 222540068 },
                new TimeZoneSummary { MicrosoftIndex = 260, Id ="E. Australia Standard Time", FullName = "(GMT+10:00) Brisbane", OptionSetValue = 222540069 },
                new TimeZoneSummary { MicrosoftIndex = 265, Id ="Tasmania Standard Time", FullName = "(GMT+10:00) Hobart", OptionSetValue = 222540070 },
                new TimeZoneSummary { MicrosoftIndex = 270, Id ="Vladivostok Standard Time", FullName = "(GMT+10:00) Vladivostok", OptionSetValue = 222540071 },
                new TimeZoneSummary { MicrosoftIndex = 275, Id ="West Pacific Standard Time", FullName = "(GMT+10:00) Guam, Port Moresby", OptionSetValue = 222540072 },
                new TimeZoneSummary { MicrosoftIndex = 280, Id ="Central Pacific Standard Time", FullName = "(GMT+11:00) Magadan, Solomon Islands, New Caledonia", OptionSetValue = 222540073 },
                new TimeZoneSummary { MicrosoftIndex = 285, Id ="Fiji Islands Standard Time", FullName = "(GMT+12:00) Fiji Islands, Kamchatka, Marshall Islands", OptionSetValue = 222540074 },
                new TimeZoneSummary { MicrosoftIndex = 290, Id ="New Zealand Standard Time", FullName = "(GMT+12:00) Auckland, Wellington", OptionSetValue = 222540075 },
                new TimeZoneSummary { MicrosoftIndex = 300, Id ="Tonga Standard Time", FullName = "(GMT+13:00) Nuku'alofa", OptionSetValue = 222540076 },
                new TimeZoneSummary { MicrosoftIndex = 92, Id ="UTC", FullName = "(GMT) Coordinated Universal Time", OptionSetValue = 222540077 }
            };
        }

        public int MicrosoftIndex { get; set; }
        public int OptionSetValue { get; set; }
        public string Id { get; set; }
        public string FullName { get; set; }
    }
}
