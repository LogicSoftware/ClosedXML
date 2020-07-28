﻿using ClosedXML.Excel.Caching;
using System;

namespace ClosedXML.Excel
{
    internal class XLStyleValue
    {
        private static readonly XLStyleRepository Repository = new XLStyleRepository(key => new XLStyleValue(key));

        public static XLStyleValue FromKey(XLStyleKey key)
        {
            return Repository.GetOrCreate(key);
        }

        internal static readonly XLStyleValue Default = FromKey(new XLStyleKey
        {
            Alignment = XLAlignmentValue.Default.Key,
            Border = XLBorderValue.Default.Key,
            Fill = XLFillValue.Default.Key,
            Font = XLFontValue.Default.Key,
            IncludeQuotePrefix = false,
            NumberFormat = XLNumberFormatValue.Default.Key,
            Protection = XLProtectionValue.Default.Key
        });

        public XLStyleKey Key { get; private set; }

        public XLAlignmentValue Alignment { get; private set; }

        public XLBorderValue Border { get; private set; }

        public XLFillValue Fill { get; private set; }

        public XLFontValue Font { get; private set; }

        public Boolean IncludeQuotePrefix { get; private set; }

        public XLNumberFormatValue NumberFormat { get; private set; }

        public XLProtectionValue Protection { get; private set; }

        public string Name { get; set; }

        internal XLStyleValue(XLStyleKey key)
        {
            Key = key;
            Alignment = XLAlignmentValue.FromKey(Key.Alignment);
            Border = XLBorderValue.FromKey(Key.Border);
            Fill = XLFillValue.FromKey(Key.Fill);
            Font = XLFontValue.FromKey(Key.Font);
            IncludeQuotePrefix = key.IncludeQuotePrefix;
            NumberFormat = XLNumberFormatValue.FromKey(Key.NumberFormat);
            Protection = XLProtectionValue.FromKey(Key.Protection);
            Name = key.Name;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;

            var cached = obj as XLStyleValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            if (_hashCode.HasValue)
                return _hashCode.Value;

            _hashCode = -280332839 + Key.GetHashCode();
            return _hashCode.Value;
        }

        public static bool operator ==(XLStyleValue left, XLStyleValue right)
        {
            if (ReferenceEquals(left, right))
                return true;
            if (ReferenceEquals(left, null) && ReferenceEquals(right, null))
                return true;
            if (ReferenceEquals(left, null) || ReferenceEquals(right, null))
                return false;
            if (left._hashCode.HasValue && right._hashCode.HasValue &&
                left._hashCode != right._hashCode)
                return false;
            return left.Key.Equals(right.Key);
        }

        public static bool operator !=(XLStyleValue left, XLStyleValue right)
        {
            return !(left == right);
        }

        private int? _hashCode;
    }
}
