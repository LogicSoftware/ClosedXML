namespace ClosedXML.Excel.Style
{
    public class XLNamedStyle
    {
        public string Name { get; set; }

        public int? BuiltIn { get; set; }

        internal XLStyleKey StyleKey { get; set; }

        internal XLNamedStyle(XLStyleKey styleKey)
        {
            StyleKey = styleKey;
        }
    }
}
