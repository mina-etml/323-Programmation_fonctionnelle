using Gpx;

namespace Rando
{
    public partial class Rando : Form
    {
        List<TrackPoint> trackPoints = new();

        public Rando()
        {
            InitializeComponent();

            string gpxFile = @"..\..\..\gpx\Ballade_châtaignère.gpx";

            if (!File.Exists(gpxFile))
            {
                MessageBox.Show($"Fichier {gpxFile} n'est pas trouvés");
            }

            var streamReader = new StreamReader(gpxFile);


            using (GpxReader reader = new GpxReader(streamReader.BaseStream))
            {
                while (reader.Read())
                {
                    switch (reader.ObjectType)
                    {
                        case GpxObjectType.Track:
                        var gpxPoints = reader.Track.ToGpxPoints();

                            //TODO : convrtir les gpxPoints en points avec un SELECT
                            var converted = gpxPoints.
                                Select(p => new TrackPoint
                                {
                                    Elevation = p.Elevation,
                                    Latitude = p.Latitude * 10000,
                                    Longitude = p.Longitude * 10000 
                                });

                            trackPoints.AddRange(converted.ToList());
                        break;
                    }
                }
            }

            RandoReduce();
        }

        private void Rando_Form_Paint(object sender, PaintEventArgs e)
        {
            Pen myPen = new Pen(Color.Red);
            myPen.Width = 2;

            var minLat = trackPoints.Min(tp => tp.Latitude);
            var maxLat = trackPoints.Max(tp => tp.Latitude);
            var rangeLat = maxLat - minLat;
            var ratioLat = rangeLat / Width;

            var minLong = trackPoints.Min(tp => tp.Longitude);
            var maxLong = trackPoints.Max(tp => tp.Longitude);
            var rangeLong = maxLong - minLong;
            var ratioLong = rangeLong / Height;

            //Trackpoint vers points
            Point[] points = trackPoints.Select(t => new Point()
            {
                X = Convert.ToInt32((t.Latitude - minLat) * ratioLat),
                Y = Convert.ToInt32((t.Longitude - minLong) * ratioLong)
            }). ToArray();

            this.CreateGraphics().DrawLines(myPen, points);
        }

        private void RandoReduce()
        {
            //calcul de la longueur du tracé
            var distance = trackPoints.Aggregate((a, b) => {
                    b.Distance = a.Distance+ a.GetDistanceFrom(b);
                    return b;
                });

            //calcul du dénivelé positif
            var positiv = trackPoints.Aggregate((a, b) => {
                b.PosDeniv = (a.GetElevationFrom(b) > 0) ? (a.PosDeniv + a.GetElevationFrom(b)) : (a.PosDeniv) ;
                return b;
            });

            //calcul du dénivelé negatif
            var negativ = trackPoints.Aggregate((a, b) => {
                b.NegDeniv = (a.GetElevationFrom(b) < 0) ? (a.NegDeniv + a.GetElevationFrom(b)) : (a.NegDeniv);
                return b;
            });


            MessageBox.Show(distance.Distance+"km");
            MessageBox.Show(positiv.PosDeniv + "positiv");
            MessageBox.Show(negativ.NegDeniv + "positiv");



        }
    }
}
