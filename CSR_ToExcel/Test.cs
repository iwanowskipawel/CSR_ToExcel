using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSR_ToExcel
{
    class Test
    {
        public string Info { get; set; }

        public string Day { get; set; }
        public string Time { get; set; }
        public string OriginalFileName { get; set; }
        public float TireTemperature { get; set; }
        public float TirePressure { get; set; }
        public bool IsWaterOn { get; set; }
        public int Speed { get; set; }
        public short AccelerateDistance { get; set; }
        public short TestDistance { get; set; }
        public double[] FrictionResults {
            get { return _frictionResults; }
            set
            {
                _frictionResults = value;

                double avarage = 0;
                foreach (double res in _frictionResults)
                    avarage += res;
                FrictionAvarage = avarage / FrictionResults.Length;
            }
        }
        double[] _frictionResults;

        public double FrictionAvarage { get; private set; }

        public Test(string day, string originalFileName, float tireTemperature,
            float tirePressure, bool isWaterOn, byte speed, short accelerateDistance,
            short testDistance, double[] frictionResults)
        {
            Day = day;
            OriginalFileName = originalFileName;
            TirePressure = tirePressure;
            TireTemperature = tireTemperature;
            IsWaterOn = isWaterOn;
            Speed = speed;
            AccelerateDistance = accelerateDistance;
            TestDistance = testDistance;
            FrictionResults = frictionResults;
        }

        public Test() { }
    }
}
