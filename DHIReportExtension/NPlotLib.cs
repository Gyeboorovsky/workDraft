using System;
using System.Drawing;
using System.IO;
using NPlot.Bitmap;

namespace DHIReportExtension
{
    public static class NPlotLib
    {
        private static Random rand = new Random(0);
        
        public static string Plot(int width, int height)
        {
            var path = @"C:\Users\togi\Documents\nplot-console-quickstart.jpeg";
            
            var linePlot = new NPlot.LinePlot { DataSource = RandomWalk() };
            var surface = new NPlot.Bitmap.PlotSurface2D(width, height);
            surface.BackColor = Color.White;
            surface.Add(linePlot);
            surface.Title = $"Scatter Plot from a Console Application";
            surface.YAxis1.Label = "Vertical Axis Label";
            surface.XAxis1.Label = "Horizontal Axis Label";
            surface.Refresh();
            
            if (File.Exists(path))
                File.Delete(path);
            
            surface.Bitmap.Save(path);

            return path;
        }
        
        private static double[] RandomWalk(int points = 500, double start = 100, double mult = 5)
        {
            // return an array of difting random numbers
            double[] values = new double[points];
            values[0] = start;
            for (int i = 1; i < points; i++)
                values[i] = values[i - 1] + (rand.NextDouble() - .5) * mult;
            return values;
        }
    }
}