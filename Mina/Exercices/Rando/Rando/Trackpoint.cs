using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rando
{
    public class TrackPoint
    {
        private double _latitude;
        private double _longitude;
        private double? _elevation;

        public double Latitude { get => _latitude; set => _latitude = value; }
        public double Longitude { get => _longitude; set => _longitude = value; }
        public double? Elevation { get => _elevation; set => _elevation = value; }
    }
}
