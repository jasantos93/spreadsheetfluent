using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFluent.Internal
{
    internal struct PropertyLocation
    {
        public PropertyLocation(PropertyLocationSectionType type, int propertyHashCode, int row, int column)
        {
            Type = type;
            PropertyHashCode = propertyHashCode;
            Row = row;
            Column = column;
        }
        public int PropertyHashCode { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        public PropertyLocationSectionType Type { get; set; }

        internal enum PropertyLocationSectionType
        {
            Header = 1,
            Body = 2
        }
    }

    internal sealed class BlockPropertyLocations : Dictionary<string, List<PropertyLocation>>
    {
        public void Add(string key, PropertyLocation propertyLocation)
        {
            if (this.ContainsKey(key))
            {
                this[key].Add(propertyLocation);
                return;
            }
            this.Add(key, new List<PropertyLocation> { propertyLocation });
        }

        public void AddRange(string key, List<PropertyLocation> propertyLocations)
        {
            if (this.ContainsKey(key))
            {
                this[key].AddRange(propertyLocations);
                return;
            }
            this.Add(key, propertyLocations);
        }
    }
}
