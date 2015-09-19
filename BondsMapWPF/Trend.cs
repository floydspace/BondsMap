using System;
using System.Collections.Generic;
using System.Linq;

namespace BondsMapWPF
{
    struct Coordinates
    {
        public double X { get; set; }
        public double Y { get; set; }

        public Coordinates(double x, double y) : this()
        {
            X = x;
            Y = y;
        }
    }

    class Trend
    {
        int[] ArrayX;
        double[] ArrayY;
        Type TT;

        public enum Type
        { Linear, Logarithmic }

        public Trend(int[] arrayX, double[] arrayY, Type tt = Type.Linear)
        {
            ArrayX = arrayX;
            ArrayY = arrayY;
            TT = tt;
        }
        public Trend(List<int> listX, List<double> listY, Type tt = Type.Linear) :
            this(listX.ToArray(), listY.ToArray(), tt) { }

        public Trend(Dictionary<int,double> dictXY, Type tt = Type.Linear)
        {
            Array.Resize(ref ArrayX, dictXY.Count);
            Array.Resize(ref ArrayY, dictXY.Count);
            dictXY.Keys.CopyTo(ArrayX, 0);
            dictXY.Values.CopyTo(ArrayY, 0);
            TT = tt;
        }

        double AverageX()
        {
            return ArrayX.Average(t => TT == Type.Logarithmic ? Math.Log(t) : t);
        }

        double AverageY()
        {
            return ArrayY.Average();
        }

        double FactorM()
        {
            double avrX = AverageX();
            double avrY = AverageY();

            double numerator = 0, denominator = 0;
            for (int i = 0; i < ArrayX.Length; i++)
            {
                double curX = TT == Type.Logarithmic ? Math.Log(ArrayX[i]) : ArrayX[i];
                double curY = ArrayY[i];
                numerator += (curY - avrY) * (curX - avrX);
                denominator += (curX - avrX) * (curX - avrX);
            }
            return numerator / denominator;
        }

        double FactorB()
        {
            return AverageY() - FactorM() * AverageX();
        }

        public double Y(double x)
        {
            return FactorM() * (TT == Type.Logarithmic ? Math.Log(x) : x) + FactorB();
        }

        public double X(double y)
        {
            return TT == Type.Logarithmic ? Math.Exp((y - FactorB()) / FactorM()) : (y - FactorB()) / FactorM();
        }
    }
}
