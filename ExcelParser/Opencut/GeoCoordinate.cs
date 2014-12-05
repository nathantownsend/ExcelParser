using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Opencut
{
    /// <summary>
    /// Represents a geographic coordinate (latitude / longitude)
    /// </summary>
    public struct GeoCoordinate
    {
        private readonly double latitude;
        private readonly double longitude;
        private readonly string mapid;
        private readonly string desc;

        public double Latitude { get { return latitude; } }
        public double Longitude { get { return longitude; } }
        public string MapId { get { return mapid; } }
        public string Description { get { return desc; } }



        public GeoCoordinate(string latitude, string longitude, string mapId, string description) : this(Convert.ToDouble(latitude), Convert.ToDouble(longitude), mapId, description) { }
        

        public GeoCoordinate(double latitude, double longitude, string mapId, string description)
        {
            this.latitude = latitude;
            this.longitude = longitude;
            this.mapid = mapId;
            this.desc = description;
        }

        public override string ToString()
        {
            return string.Format("{0},{1},{2}", MapId, Latitude, Longitude);
        }
    }
}
