using Showy.Purcell.Model;
using System.Drawing;

namespace Showy.Purcell;

/// <summary>
/// 预设的表格样式 Preset table styles
/// </summary>
public static class TablePresetStyles
{
    /// <summary>
    /// 默认 Default
    /// </summary>
    public static readonly CustomStyle Default = new CustomStyle();

    /// <summary>
    /// 明亮清新蓝 BrightFresh
    /// 背景色名称: Sky Blue
    /// 背景色值: #00BFFF
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>，  
    public static readonly TableStyle BrightFresh =
        new TableStyle(Color.FromArgb(0, 191, 255), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 大地色调 EarthTones
    /// 背景色名称: Gray
    /// 背景色值: #808080
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle EarthTones =
        new TableStyle(Color.FromArgb(128, 128, 128), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 暖色调 WarmTones
    /// 背景色名称: Red
    /// 背景色值: #FF0000
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle
        WarmTones = new TableStyle(Color.FromArgb(255, 0, 0), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 海洋蓝 OceanBlue
    /// 背景色名称: Midnight Blue
    /// 背景色值: #191970
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle OceanBlue =
        new TableStyle(Color.FromArgb(25, 25, 112), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 复古怀旧 VintageNostalgia
    /// 背景色名称: Pink
    /// 背景色值: #FFC0CB
    /// 文本色名称: Gray
    /// 文本色值: #808080
    /// </summary>
    public static readonly TableStyle VintageNostalgia =
        new TableStyle(Color.FromArgb(255, 192, 203), Color.FromArgb(128, 128, 128), 10, 10, 10, 10);

    /// <summary>
    /// 极简黑白 MinimalistBW
    /// 背景色名称: White
    /// 背景色值: #FFFFFF
    /// 文本色名称: Gray
    /// 文本色值: #808080
    /// </summary>
    public static readonly TableStyle MinimalistBw =
        new TableStyle(Color.FromArgb(255, 255, 255), Color.FromArgb(128, 128, 128), 10, 10, 10, 10);

    /// <summary>
    /// 活力能量 VibrantEnergy
    /// 背景色名称: Orange
    /// 背景色值: #FFA500
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle VibrantEnergy =
        new TableStyle(Color.FromArgb(255, 165, 0), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 复古时尚 RetroChic
    /// 背景色名称: Orchid
    /// 背景色值: #DA70D6
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle RetroChic =
        new TableStyle(Color.FromArgb(218, 112, 214), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 温馨秋日 CozyAutumn
    /// 背景色名称: Peru
    /// 背景色值: #CD853F
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle CozyAutumn =
        new TableStyle(Color.FromArgb(205, 133, 63), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 宁静自然 SereneNature
    /// 背景色名称: Sea Green
    /// 背景色值: #2E8B57
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle SereneNature =
        new TableStyle(Color.FromArgb(46, 139, 87), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 午夜魔幻 MidnightMagic
    /// 背景色名称: Navy
    /// 背景色值: #000080
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle MidnightMagic =
        new TableStyle(Color.FromArgb(0, 0, 128), Color.White, 10, 10, 10, 10);

    /// <summary>
    /// 暖阳阳光 SunnyDay
    /// 背景色名称: Yellow
    /// 背景色值: #FFFF00
    /// 文本色名称: Gray
    /// 文本色值: #808080
    /// </summary>
    public static readonly TableStyle SunnyDay =
        new TableStyle(Color.FromArgb(255, 255, 0), Color.FromArgb(128, 128, 128), 10, 10, 10, 10);

    /// <summary>
    /// 优雅单色 ElegantMonochrome
    /// 背景色名称: Dark Gray
    /// 背景色值: #A9A9A9
    /// 文本色名称: White
    /// 文本色值: #FFFFFF
    /// </summary>
    public static readonly TableStyle ElegantMonochrome =
        new TableStyle(Color.FromArgb(169, 169, 169), Color.White, 10, 10, 10, 10);
}