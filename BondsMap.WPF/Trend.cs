using System;
using System.Collections.Generic;
using System.Linq;

namespace BondsMapWPF
{
    struct Point
    {
        public double X { get; private set; }
        public double Y { get; private set; }

        public Point(double x, double y) : this()
        {
            X = x;
            Y = y;
        }
    }

    class Trend
    {
        private readonly Type _tt;
        private readonly Point[] _points;

        public enum Type
        { Linear, Logarithmic }

        public Trend(Point[] points, Type tt = Type.Linear)
        {
            _points = points;
            _tt = tt;
        }

        private double AverageX
        {
            get { return _points.Average(p => _tt == Type.Logarithmic ? Math.Log(p.X) : p.X); }
        }

        private double AverageY
        {
            get { return _points.Average(p => p.Y); }
        }

        public double FactorM
        {
            get
            {
                double numerator = 0, denominator = 0;
                foreach (var point in _points)
                {
                    double curX = _tt == Type.Logarithmic ? Math.Log(point.X) : point.X;
                    double curY = point.Y;
                    numerator += (curY - AverageY)*(curX - AverageX);
                    denominator += (curX - AverageX)*(curX - AverageX);
                }
                return numerator/denominator;
            }
        }

        public double FactorB
        {
            get { return AverageY - FactorM*AverageX; }
        }

        public double Y(double x)
        {
            return FactorM * (_tt == Type.Logarithmic ? Math.Log(x) : x) + FactorB;
        }

        public double X(double y)
        {
            return _tt == Type.Logarithmic ? Math.Exp((y - FactorB) / FactorM) : (y - FactorB) / FactorM;
        }
    }
}
