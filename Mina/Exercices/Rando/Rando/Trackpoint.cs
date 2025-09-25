using Gpx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rando
{
    public class TrackPoint
    {
        private const double EARTH_RADIUS = 6371; // [km]
        private const double RADIAN = Math.PI / 180;

        private double _latitude;
        private double _longitude;
        private double? _elevation;

        public double Latitude { get => _latitude; set => _latitude = value; }
        public double Longitude { get => _longitude; set => _longitude = value; }
        public double? Elevation { get => _elevation; set => _elevation = value; }

        public double Distance { get; set; } = 0;
        public double PosDeniv { get; set; } = 0;
        public double NegDeniv { get; set; } = 0;


        //Decimetre
        public double GetDistanceFrom(TrackPoint other)
        {
            double thisLatitude = Latitude * RADIAN;
            double otherLatitude = other.Latitude * RADIAN;
            double deltaLongitude = Math.Abs(Longitude - other.Longitude) * RADIAN;

            double cos = Math.Cos(deltaLongitude) * Math.Cos(thisLatitude) * Math.Cos(otherLatitude) +
                Math.Sin(thisLatitude) * Math.Sin(otherLatitude);

            double distance = EARTH_RADIUS * Math.Acos(Math.Max(Math.Min(cos, 1), -1)); //decimeter
            double km = distance / 10000; // to km

            return km;
        }

        public double GetElevationFrom(TrackPoint other)
        {
            if (!this.Elevation.HasValue || !other.Elevation.HasValue)
                return 0;

            return other.Elevation.Value - this.Elevation.Value;
        }
    }
}
