using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace ClosedXML.Excel
{
    internal class XLTheme : IXLTheme
    {
        public XLColor Background1 { get; set; }
        public XLColor Text1 { get; set; }
        public XLColor Background2 { get; set; }
        public XLColor Text2 { get; set; }
        public XLColor Accent1 { get; set; }
        public XLColor Accent2 { get; set; }
        public XLColor Accent3 { get; set; }
        public XLColor Accent4 { get; set; }
        public XLColor Accent5 { get; set; }
        public XLColor Accent6 { get; set; }
        public XLColor Hyperlink { get; set; }
        public XLColor FollowedHyperlink { get; set; }

        public XLTheme()
        {}

        public XLTheme(ThemePart themePart)
        {
            static XLColor Func(Color2Type color)
            {
                return color.SystemColor != null ? XLColor.FromHtml("#FF" + color.SystemColor.LastColor) : color.RgbColorModelHex != null ? XLColor.FromHtml("#FF" + color.RgbColorModelHex.Val) : XLColor.FromHtml("#FF000000");
            }

            Text1 = Func(themePart.Theme.ThemeElements.ColorScheme.Dark1Color);
            Background1 = Func(themePart.Theme.ThemeElements.ColorScheme.Light2Color);
            Text2 = Func(themePart.Theme.ThemeElements.ColorScheme.Light1Color);
            Background2 = Func(themePart.Theme.ThemeElements.ColorScheme.Dark2Color);
            Accent1 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent1Color);
            Accent2 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent2Color);
            Accent3 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent3Color);
            Accent4 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent4Color);
            Accent5 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent5Color);
            Accent6 = Func(themePart.Theme.ThemeElements.ColorScheme.Accent6Color);
            Hyperlink = Func(themePart.Theme.ThemeElements.ColorScheme.Hyperlink);
            FollowedHyperlink = Func(themePart.Theme.ThemeElements.ColorScheme.FollowedHyperlinkColor);
        }

        public XLColor ResolveThemeColor(XLThemeColor themeColor)
        {
            return themeColor switch
            {
                XLThemeColor.Background1 => Background1,
                XLThemeColor.Text1 => Text1,
                XLThemeColor.Background2 => Background2,
                XLThemeColor.Text2 => Text2,
                XLThemeColor.Accent1 => Accent1,
                XLThemeColor.Accent2 => Accent2,
                XLThemeColor.Accent3 => Accent3,
                XLThemeColor.Accent4 => Accent4,
                XLThemeColor.Accent5 => Accent5,
                XLThemeColor.Accent6 => Accent6,
                XLThemeColor.Hyperlink => Hyperlink,
                XLThemeColor.FollowedHyperlink => FollowedHyperlink,
                _ => null,
            };
        }
    }
}
