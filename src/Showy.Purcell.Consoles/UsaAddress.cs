using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static LargeXlsx.XlsxFill;

namespace Showy.Purcell.Consoles;

internal static class UsaAddress
{
    internal static void Proc()
    {
        System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
        sw.Start();
        string file = @"D:\MOD-houzz数据0531.xlsx";

        int counts1 = 0;
        int counts2 = 0;

        Dictionary<string, UsaRegionWithShort> ursSet = new Dictionary<string, UsaRegionWithShort>();
        foreach (var item in usaRegionWithShorts)
        {
            ursSet[item.ProvinceShort ?? string.Empty] = item;
        }

        var list = new List<HouzzData>();
        foreach (var item in Purcell.Query(file))
        {
            var data = new HouzzData
            {
                BusinessName = item["BusinessName"]?.ToString(),
                PhoneNumber = item["PhoneNumber"]?.ToString(),
                Email = item["Email"]?.ToString(),
                License = item["License"]?.ToString(),
                Website = item["Website"]?.ToString(),
                Address = item["Address"]?.ToString(),
                Url = item["Url"]?.ToString(),
                Directory = item["Directory"]?.ToString(),
                Socials = item["Socials"]?.ToString(),
            };
            if (!string.IsNullOrWhiteSpace(data.Address))
            {
                // 澳大利亚 206 Princes HighwaySylvania Heights, New South Wales 2224Australia
                // ,\s+(([A-Z][a-z]+\s+)+)(\d{4})([A-Z][a-z]+)
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat5 = Regex.Match(data.Address, @",\s+(([A-Z][a-z]+\s+)+)(\d{4})(Australia)\b");
                    if (mat5.Success)
                    {
                        data.Province = mat5.Groups[1].Value;
                        data.ZipCode = mat5.Groups[mat5.Groups.Count - 2].Value;
                        data.Country = mat5.Groups[mat5.Groups.Count - 1].Value;
                    }
                }
                // 加拿大 116 172ndSurrey, British Columbia V3S 9R2Canada
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat2 = Regex.Match(data.Address, @"\b([A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d)(Canada)$");
                    if (mat2.Success)
                    {
                        data.ZipCode = mat2.Groups[1].Value;
                        data.Country = mat2.Groups[mat2.Groups.Count - 1].Value;
                    }
                }
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat2 = Regex.Match(data.Address, @"\b([A-Za-z]\d[A-Za-z][ -]?\d[A-Za-z]\d?)(Canada)$");
                    if (mat2.Success)
                    {
                        data.ZipCode = mat2.Groups[1].Value;
                        data.Country = mat2.Groups[mat2.Groups.Count - 1].Value;
                    }
                }
                // 西班牙
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat7 = Regex.Match(data.Address, @"(\d{5})(([A-Z][a-z]+.+)+)(Spain)$");
                    if (mat7.Success)
                    {
                        data.ZipCode = mat7.Groups[1].Value;
                        data.Country = mat7.Groups[mat7.Groups.Count - 1].Value;
                    }
                }
                // 德国 Pappelallee 5710437 BerlinGermany
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat8 = Regex.Match(data.Address, @"(\d{5}) (([A-Z][a-z]+)+)(Germany)$");
                    if (mat8.Success)
                    {
                        data.ZipCode = mat8.Groups[1].Value;
                        data.Country = mat8.Groups[mat8.Groups.Count - 1].Value;
                    }
                }
                // 瑞典 116 34StockholmSweden
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat9 = Regex.Match(data.Address, @"(\d{3} \d{2})(([A-Z][a-z]+)+)(Sweden)$");
                    if (mat9.Success)
                    {
                        data.ZipCode = mat9.Groups[1].Value;
                        data.Province = mat9.Groups[mat9.Groups.Count - 2].Value;
                        data.Country = mat9.Groups[mat9.Groups.Count - 1].Value;
                    }
                }
                // 法国 87700Aixe sur VienneFrance
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat10 = Regex.Match(data.Address, @"(\d{5})(([A-Z][a-z\s]+)+)(France)$");
                    if (mat10.Success)
                    {
                        data.ZipCode = mat10.Groups[1].Value;
                        data.Country = mat10.Groups[mat10.Groups.Count - 1].Value;
                    }
                }
                // 新加坡 33 Ubi Ave 3, Vertex Tower A#05-72 SSingaporeSingapore408868
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat10 = Regex.Match(data.Address, @"(Singapore)(\d{6})$");
                    if (mat10.Success)
                    {
                        data.Country = mat10.Groups[1].Value;
                        data.ZipCode = mat10.Groups[mat10.Groups.Count - 1].Value;
                    }
                }
                // 印第安：\b([A-Z][a-z]+)([A-Z][a-z]+)(India)(\d{6})\b
                // No-45, 7th Cross, 16 B Main, 4th B Block, Koramangala Extension,Behind Koramangala B.D.A Complex,BengaluruKarnatakaIndia560034
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat3 = Regex.Match(data.Address, @"\b([A-Z][a-z]+)([A-Z][a-z]+)(India)(\d{6})\b");
                    if (mat3.Success)
                    {
                        data.City = mat3.Groups[1].Value;
                        data.Province = mat3.Groups[2].Value;
                        data.Country = mat3.Groups[3].Value;
                        data.ZipCode = mat3.Groups[4].Value;
                    }
                }
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat3v1 = Regex.Match(data.Address, @"(India)(\d{6})$");
                    if (mat3v1.Success)
                    {
                        data.Country = mat3v1.Groups[1].Value;
                        data.ZipCode = mat3v1.Groups[2].Value;
                    }
                }
                // 美国 3661 North Campbell Avenue PMB 208Tucson, AZ 85719
                // ,\s+([A-Z]{2})\s{0,2}(\d{5})
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat6 = Regex.Match(data.Address, @",\s+([A-Z]{2})\s{0,2}(\d{5})");
                    if (mat6.Success)
                    {
                        data.ProvinceShort = mat6.Groups[1].Value;
                        data.ZipCode = mat6.Groups[2].Value;
                        data.Country = "United States";
                    }
                }
                // 美国短 Houston, TX
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat6v1 = Regex.Match(data.Address, @"([A-Z]?[a-z]+),\s{0,2}([A-Z]{2})$");
                    if (mat6v1.Success)
                    {
                        data.City = mat6v1.Groups[1].Value;
                        data.ProvinceShort = mat6v1.Groups[mat6v1.Groups.Count - 1].Value;
                        data.Country = "United States";
                    }
                }
                // 美国短 Houston, TX
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat6v2 = Regex.Match(data.Address, @",\s{1}([A-Z]{2})$");
                    if (mat6v2.Success)
                    {
                        data.ProvinceShort = mat6v2.Groups[1].Value;
                        data.Country = "United States";
                    }
                }
                // 美国短 , TX
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat6v1 = Regex.Match(data.Address, @",\s{1}([A-Z]{2})$");
                    if (mat6v1.Success)
                    {
                        data.ProvinceShort = mat6v1.Groups[mat6v1.Groups.Count - 1].Value;
                        data.Country = "United States";
                    }
                }
                // 美国夏威夷 Kailua Kona 77-6465 Princess Keelikolani DrKailua Kona 96740
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat6v1 = Regex.Match(data.Address, @"Kailua Kona (\d{5})$");
                    if (mat6v1.Success)
                    {
                        data.ZipCode = mat6v1.Groups[mat6v1.Groups.Count - 1].Value;
                        data.ProvinceShort = "HI";
                        data.Country = "United States";
                    }
                }
                // 英国 (([A-Z]{1,2}[0-9][A-Z0-9]?)\s*([0-9][A-Z]{2}))
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat4 = Regex.Match(data.Address, @"\b(([A-Z]{1,2}[0-9][A-Z0-9]?)\s*([0-9][A-Z]{2}))([A-Za-z][A-Za-z\s]+)\b");
                    if (mat4.Success)
                    {
                        data.ZipCode = mat4.Groups[1].Value;
                        data.Country = mat4.Groups[mat4.Groups.Count - 1].Value;
                    }
                }
                // 意大利 via Federico Confalonieri 9/A95123 CataniaItaly
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat10 = Regex.Match(data.Address, @"(\d{5}) ([A-Za-z]+)+(Italy)$");
                    if (mat10.Success)
                    {
                        data.ZipCode = mat10.Groups[1].Value;
                        data.Country = mat10.Groups[mat10.Groups.Count - 1].Value;
                    }
                }
                // 新西兰 Level 2/60 College HillFreemansAuckland1010New Zealand
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat10 = Regex.Match(data.Address, @"(\d{4})(New Zealand)$");
                    if (mat10.Success)
                    {
                        data.ZipCode = mat10.Groups[1].Value;
                        data.Country = mat10.Groups[mat10.Groups.Count - 1].Value;
                    }
                }
                // 匹配邮编和国家结尾
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    Match mat1 = Regex.Match(data.Address, @"\b(\d{5,6})([A-Z][A-Za-z\s]+)$");
                    if (mat1.Success)
                    {
                        data.ZipCode = mat1.Groups[1].Value;
                        data.Country = mat1.Groups[mat1.Groups.Count - 1].Value;
                    }
                }
                if (string.IsNullOrWhiteSpace(data.Country))
                {
                    if (data.Address.EndsWith("United States") || data.Address.EndsWith("UnitedStates"))
                        data.Country = "United States";
                    else if (data.Address.EndsWith("United Kingdom") || data.Address.EndsWith("UnitedKingdom"))
                        data.Country = "United Kingdom";
                    else if (data.Address.EndsWith("Germany"))
                        data.Country = "Germany";
                    else if (data.Address.EndsWith("Canada"))
                        data.Country = "Canada";
                    else if (data.Address.EndsWith("Sweden"))
                        data.Country = "Sweden";
                    else if (data.Address.EndsWith("Spain"))
                        data.Country = "Spain";
                    else if (data.Address.EndsWith("Australia"))
                        data.Country = "Australia";
                    else if (data.Address.EndsWith("France"))
                        data.Country = "France";
                    else if (data.Address.EndsWith("Italy"))
                        data.Country = "Italy";
                    else if (data.Address.EndsWith("New Zealand"))
                        data.Country = "New Zealand";
                    else if (data.Address.EndsWith("Singapore"))
                        data.Country = "Singapore";
                    else if (data.Address.Contains("香港")
                        || data.Address.Trim().StartsWith("hk", StringComparison.OrdinalIgnoreCase)
                        || data.Address.Trim().EndsWith("hk", StringComparison.OrdinalIgnoreCase)
                        || data.Address.Contains("H.K."))
                        data.Country = "Hong Kong";
                }
                if (data.Country != null && data.Country.EndsWith("France") && !string.IsNullOrWhiteSpace(data.ZipCode))
                {
                    data.Country = "France";
                }
                if (data.Country != null && data.Country.Equals("london"))
                {
                    data.Country = "London";
                }
                if (!string.IsNullOrWhiteSpace(data.ProvinceShort))
                {
                    var itemUrs = !ursSet.ContainsKey(data.ProvinceShort) ? null : ursSet[data.ProvinceShort];

                    if (itemUrs != null)
                    {
                        data.Province = itemUrs.Province;
                        data.Region = itemUrs.Region;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(data.BusinessName) && !string.IsNullOrWhiteSpace(data.Address))
                counts1++;
            if (!string.IsNullOrWhiteSpace(data.ZipCode) && !string.IsNullOrWhiteSpace(data.Country))
                counts2++;

            list.Add(data);
        }
        Console.WriteLine($"有效数据量: {counts1}");
        Console.WriteLine($"解析邮编国家 格式1 的量: {counts2}");
        Console.WriteLine($"解析邮编国家 有国 的量: {list.Where(q => !string.IsNullOrWhiteSpace(q.Country)).Count()}");
        var next = list.Where(q => !string.IsNullOrWhiteSpace(q.Address) && string.IsNullOrWhiteSpace(q.Country))
            .Skip(4).FirstOrDefault()?.Address ?? string.Empty;
        Console.WriteLine($"未解析下一格式：{next}");

        //Purcell.Export(@"d:\国外地址.xlsx", list);
        sw.Stop();
        Console.WriteLine($"解析完成，耗时: {sw.ElapsedMilliseconds} ms");
    }

    #region 美国洲和短写

    internal static List<UsaRegionWithShort> usaRegionWithShorts = [
        new UsaRegionWithShort { Region = "New England", Province = "Connecticut", ProvinceShort = "CT" },
        new UsaRegionWithShort { Region = "New England", Province = "Maine", ProvinceShort = "ME" },
        new UsaRegionWithShort { Region = "New England", Province = "Massachusetts", ProvinceShort = "MA" },
        new UsaRegionWithShort { Region = "New England", Province = "New Hampshire", ProvinceShort = "NH" },
        new UsaRegionWithShort { Region = "New England", Province = "Rhode Island", ProvinceShort = "RI" },
        new UsaRegionWithShort { Region = "New England", Province = "Vermont", ProvinceShort = "VT" },
        new UsaRegionWithShort { Region = "Mideast", Province = "Delaware", ProvinceShort = "DE" },
        new UsaRegionWithShort { Region = "Mideast", Province = "District of Columbia", ProvinceShort = "DC" },
        new UsaRegionWithShort { Region = "Mideast", Province = "Maryland", ProvinceShort = "MD" },
        new UsaRegionWithShort { Region = "Mideast", Province = "New Jersey", ProvinceShort = "NJ" },
        new UsaRegionWithShort { Region = "Mideast", Province = "New York", ProvinceShort = "NY" },
        new UsaRegionWithShort { Region = "Mideast", Province = "Pennsylvania", ProvinceShort = "PA" },
        new UsaRegionWithShort { Region = "Great Lakes", Province = "Illinois", ProvinceShort = "IL" },
        new UsaRegionWithShort { Region = "Great Lakes", Province = "Indiana", ProvinceShort = "IN" },
        new UsaRegionWithShort { Region = "Great Lakes", Province = "Michigan", ProvinceShort = "MI" },
        new UsaRegionWithShort { Region = "Great Lakes", Province = "Ohio", ProvinceShort = "OH" },
        new UsaRegionWithShort { Region = "Great Lakes", Province = "Wisconsin", ProvinceShort = "WI" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Alabama", ProvinceShort = "AL" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Arkansas", ProvinceShort = "AR" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Florida", ProvinceShort="FL"},
        new UsaRegionWithShort { Region = "Southeast", Province = "Georgia", ProvinceShort = "GA" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Kentucky", ProvinceShort = "KY" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Louisiana", ProvinceShort="LA"},
        new UsaRegionWithShort { Region = "Southeast", Province = "Mississippi", ProvinceShort = "MS" },
        new UsaRegionWithShort { Region = "Southeast", Province = "North Carolina", ProvinceShort = "NC" },
        new UsaRegionWithShort { Region = "Southeast", Province = "South Carolina", ProvinceShort="SC"},
        new UsaRegionWithShort { Region = "Southeast", Province = "Tennessee", ProvinceShort = "TN" },
        new UsaRegionWithShort { Region = "Southeast", Province = "Virginia", ProvinceShort = "VA" },
        new UsaRegionWithShort { Region = "Southeast", Province = "West Virginia", ProvinceShort="WV"},
        new UsaRegionWithShort { Region = "Southwest", Province = "Arizona", ProvinceShort="AZ"},
        new UsaRegionWithShort { Region = "Southwest", Province = "New Mexico", ProvinceShort = "NM" },
        new UsaRegionWithShort { Region = "Southwest", Province = "Oklahoma", ProvinceShort = "OK" },
        new UsaRegionWithShort { Region = "Southwest", Province = "Texas", ProvinceShort="TX"},
        new UsaRegionWithShort { Region = "Rocky Mountain", Province = "Colorado", ProvinceShort="CO"},
        new UsaRegionWithShort { Region = "Rocky Mountain", Province = "Idaho", ProvinceShort = "ID" },
        new UsaRegionWithShort { Region = "Rocky Mountain", Province = "Montana", ProvinceShort = "MT" },
        new UsaRegionWithShort { Region = "Rocky Mountain", Province = "Utah", ProvinceShort="UT"},
        new UsaRegionWithShort { Region = "Rocky Mountain", Province = "Wyoming", ProvinceShort="WY"},
        new UsaRegionWithShort { Region = "Plains", Province = "Iowa", ProvinceShort = "IA" },
        new UsaRegionWithShort { Region = "Plains", Province = "Kansas", ProvinceShort = "KS" },
        new UsaRegionWithShort { Region = "Plains", Province = "Minnesota", ProvinceShort="MN"},
        new UsaRegionWithShort { Region = "Plains", Province = "Missouri", ProvinceShort="MO"},
        new UsaRegionWithShort { Region = "Plains", Province = "Nebraska", ProvinceShort = "NE" },
        new UsaRegionWithShort { Region = "Plains", Province = "North Dakota", ProvinceShort = "ND" },
        new UsaRegionWithShort { Region = "Plains", Province = "South Dakota", ProvinceShort="SD"},
        new UsaRegionWithShort { Region = "Far West", Province = "Alaska", ProvinceShort="AK"},
        new UsaRegionWithShort { Region = "Far West", Province = "California", ProvinceShort = "CA" },
        new UsaRegionWithShort { Region = "Far West", Province = "Hawaii", ProvinceShort = "HI" },
        new UsaRegionWithShort { Region = "Far West", Province = "Nevada", ProvinceShort="NV"},
        new UsaRegionWithShort { Region = "Far West", Province = "Oregon", ProvinceShort="OR"},
        new UsaRegionWithShort { Region = "Far West", Province = "Washington", ProvinceShort="WA"},
    ];

    #endregion 美国洲和短写

}

internal class HouzzData
{
    public string? BusinessName { get; set; }

    public string? PhoneNumber { get; set; }

    public string? Email { get; set; }

    public string? License { get; set; }

    public string? Website { get; set; }

    public string? Address { get; set; }

    public string? Url { get; set; }

    public string? Directory { get; set; }

    public string? Socials { get; set; }

    // 国家
    public string? Country { get; set; }

    // 东西南北的区域
    public string? Region { get; set; }

    // 省/洲
    public string? Province { get; set; }

    // 省/洲 缩写
    public string? ProvinceShort { get; set; }

    // 市
    public string? City { get; set; }

    // 邮编
    public string? ZipCode { get; set; }
}

internal class UsaRegionWithShort
{
    internal string? Region { get; set; }
    internal string? Province { get; set; }
    internal string? ProvinceShort { get; set; }
}