using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Showy.Purcell;

/**Below Code Copy And Modified from ClosedXML(https://github.com/ClosedXML) @MIT License **/
internal static class XmlEncoder
{
    private static readonly Regex XHhhhRegex = new("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
    private static readonly Regex UppercaseXHhhhRegex = new("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled);


    private static readonly Regex EscapeRegex = new("_x([0-9A-F]{4,4})_");

    public static string EncodeString(string? encodeStr)
    {
        if (encodeStr == null) return string.Empty;

        encodeStr = XHhhhRegex.Replace(encodeStr, "_x005F_$1_");

        var sb = new StringBuilder(encodeStr.Length);

        foreach (var ch in encodeStr)
            if (XmlConvert.IsXmlChar(ch))
                sb.Append(ch);
            else
                sb.Append(XmlConvert.EncodeName(ch.ToString()));

        return sb.ToString();
    }

    public static string? DecodeString(string? decodeStr)
    {
        if (string.IsNullOrEmpty(decodeStr))
            return string.Empty;
        decodeStr = UppercaseXHhhhRegex.Replace(decodeStr, "_x005F_$1_");
        return XmlConvert.DecodeName(decodeStr);
    }

    public static string ConvertEscapeChars(string input)
    {
        return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
    }
}