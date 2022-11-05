using System;

namespace Delaney.DocX
{
    public class Office
    {
        public static Measurment DefaultMeasurement = Measurment.Points;

        public static int CentimetersToPoints(double nCentimeters)
        {
            return Convert.ToInt32(nCentimeters * 566.95);
        }

        public static int GetPercent(double nPercent)
        {
            // Guard Clause
            if (!(nPercent > 0))
                return 0;

            if (nPercent > 100)
                nPercent = 100;

            return Convert.ToInt32(nPercent * 50);
        }
    }
}
