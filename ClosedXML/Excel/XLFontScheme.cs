using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Excel
{
    public class XLFontScheme
    {
        private class CollectionFont
        {
            private string LatinFont { get; }
            private string EastAsianFont { get; }
            private string ComplexScriptFont { get; }

            private Dictionary<string, string> SupplementalFonts { get; }

            public CollectionFont(FontCollectionType fontCollection)
            {
                LatinFont = fontCollection.LatinFont.Typeface;
                EastAsianFont = fontCollection.EastAsianFont.Typeface;
                ComplexScriptFont = fontCollection.ComplexScriptFont.Typeface;
                SupplementalFonts = fontCollection.Elements<SupplementalFont>().ToDictionary(sf => sf.Script.ToString(), sf => sf.Typeface.ToString());
            }

            public void FillExcelCollectionFont(FontCollectionType excelFont)
            {
                excelFont.AppendChild(new LatinFont { Typeface = LatinFont });
                excelFont.AppendChild(new EastAsianFont { Typeface = EastAsianFont });
                excelFont.AppendChild(new ComplexScriptFont { Typeface = ComplexScriptFont });

                SupplementalFonts.Select(pair => new SupplementalFont { Script = pair.Key, Typeface = pair.Value })
                    .ForEach(sf => excelFont.AppendChild(sf));
            }
        }

        private CollectionFont MajorFont { get; set; }
        private CollectionFont MinorFont { get; set; }
        public string Name { get; set; }

        public XLFontScheme(FontScheme fontScheme)
        {
            if (fontScheme == null)
                return;

            MajorFont = new CollectionFont(fontScheme.MajorFont);
            MinorFont = new CollectionFont(fontScheme.MinorFont);
            Name = fontScheme.Name;
        }

        public FontScheme ToExcelFontScheme()
        {
            var majorFont = new MajorFont();
            MajorFont.FillExcelCollectionFont(majorFont);

            var minorFont = new MinorFont();
            MinorFont.FillExcelCollectionFont(minorFont);

            var fontScheme = new FontScheme { Name = string.IsNullOrWhiteSpace(Name) ? "Office" : Name };

            fontScheme.AppendChild(majorFont);
            fontScheme.AppendChild(minorFont);

            return fontScheme;
        }

        public bool IsEmpty()
        {
            return MajorFont == null || MinorFont == null;
        }
    }
}
