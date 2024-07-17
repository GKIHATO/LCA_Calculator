using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelloWorld
{
    internal class UnitConverter
    {
        double ratio; // Field

        public UnitConverter(double unitRatio) // Constructor
        {
            ratio = unitRatio;
        }

        public UnitConverter(string originUnit, string targetUnit) // Constructor
        {
            (string, string) standardOriginUnit = GetStandardUnit(originUnit);

            (string, string) standardTargetUnit = GetStandardUnit(targetUnit);

            if (standardOriginUnit.Item2 == standardTargetUnit.Item2)
            {
                switch (standardOriginUnit.Item2)
                {
                    case "Volume":
                        ratio = ConversionVolume[standardOriginUnit.Item1] / ConversionVolume[standardTargetUnit.Item1];
                        break;
                    case "Area":
                        ratio = ConversionArea[standardOriginUnit.Item1] / ConversionArea[standardTargetUnit.Item1];
                        break;
                    case "Weight":
                        ratio = ConversionWeight[standardOriginUnit.Item1] / ConversionWeight[standardTargetUnit.Item1];
                        break;
                    case "Length":
                        ratio = ConversionLength[standardOriginUnit.Item1] / ConversionLength[standardTargetUnit.Item1];
                        break;
                    case "Count":
                        ratio = 1;
                        break;
                    default:
                        ratio = 0;
                        break;
                }
            }
            else
            {
                ratio = 0;
            }
        }

        public double Convert(double unit) // Method
        {
            return unit * ratio;
        }

        public (string, string) GetStandardUnit(string unit)
        {
            string unitLower = unit.ToLower().Trim();

            switch (unitLower)
            {
                case "m3":
                case "m^3":
                case "cubicmeter":
                case "cubicmeters":
                case "cubicmetre":
                case "cubicmetres":
                    return ("m3", "Volume");

                case "cm3":
                case "cm^3":
                case "cubiccentimeter":
                case "cubiccentimeters":
                case "cubiccentimetre":
                case "cubiccentimetres":
                    return ("cm3", "Volume");

                case "mm3":
                case "mm^3":
                case "cubicmillimeter":
                case "cubicmillimeters":
                case "cubicmillimetre":
                case "cubicmillimetres":
                    return ("mm3", "Volume");

                case "dm3":
                case "dm^3":
                case "cubicdecimeter":
                case "cubicdecimeters":
                case "cubicdecimetre":
                case "cubicdecimetres":
                    return ("dm3", "Volume");

                case "ft3":
                case "ft^3":
                case "cubicfoot":
                case "cubicfeet":
                    return ("ft3", "Volume");

                case "in3":
                case "in^3":
                case "cubicinch":
                case "cubicinches":
                    return ("in3", "Volume");

                case "yd3":
                case "yd^3":
                case "cubicyard":
                case "cubicyards":
                    return ("yd3", "Volume");

                case "L":
                case "liter":
                case "liters":
                case "litre":
                case "litres":
                    return ("L", "Volume");

                case "t":
                case "ton":
                case "tons":
                case "tonne":
                case "tonnes":
                    return ("ton", "Weight");

                case "kg":
                case "kilogram":
                case "kilograms":
                    return ("kg", "Weight");

                case "g":
                case "gram":
                case "grams":
                    return ("g", "Weight");

                case "lb":
                case "lbs":
                case "pound":
                case "pounds":
                    return ("lb", "Weight");

                case "m":
                case "meter":
                case "meters":
                case "metre":
                case "metres":
                    return ("m", "Length");

                case "cm":
                case "centimeter":
                case "centimeters":
                case "centimetre":
                case "centimetres":
                    return ("cm", "Length");

                case "mm":
                case "millimeter":
                case "millimeters":
                case "millimetre":
                case "millimetres":
                    return ("mm", "Length");

                case "dm":
                case "decimeter":
                case "decimeters":
                case "decimetre":
                case "decimetres":
                    return ("dm", "Length");

                case "km":
                case "kilometer":
                case "kilometers":
                case "kilometre":
                case "kilometres":
                    return ("km", "Length");

                case "ft":
                case "foot":
                case "feet":
                    return ("ft", "Length");

                case "in":
                case "inch":
                case "inches":
                    return ("in", "Length");

                case "yd":
                case "yard":
                case "yards":
                    return ("yd", "Length");

                case "m2":
                case "m^2":
                case "squaremeter":
                case "squaremeters":
                case "squaremetre":
                case "squaremetres":
                    return ("m2", "Area");

                case "cm2":
                case "cm^2":
                case "squarecentimeter":
                case "squarecentimeters":
                case "squarecentimetre":
                case "squarecentimetres":
                    return ("cm2", "Area");

                case "mm2":
                case "mm^2":
                case "squaremillimeter":
                case "squaremillimeters":
                case "squaremillimetre":
                case "squaremillimetres":
                    return ("mm2", "Area");

                case "dm2":
                case "dm^2":
                case "squaredecimeter":
                case "squaredecimeters":
                case "squaredecimetre":
                case "squaredecimetres":
                    return ("dm2", "Area");

                case "km2":
                case "km^2":
                case "squarekilometer":
                case "squarekilometers":
                    return ("km2", "Area");

                case "ft2":
                case "ft^2":
                case "squarefoot":
                case "squarefeet":
                    return ("ft2", "Area");

                case "in2":
                case "in^2":
                case "squareinch":
                case "squareinches":
                    return ("in2", "Area");

                case "yd2":
                case "yd^2":
                case "squareyard":
                case "squareyards":
                    return ("yd2", "Area");

                case "ea.":
                    return ("ea.", "Count");

                default:
                    return ("", "");

            }
        }

        /*
                Dictionary<(string, string), double> ConversionChart = new Dictionary<(string, string), double>
                {
                    {("cm3", "m3" ), 0.000001},
                    {("mm3", "m3" ), 0.000000001},
                    {("dm3", "m3" ), 0.001},
                    {("ft3", "m3" ), 0.0283168},
                    {("in3", "m3" ), 0.0000163871},
                    {("yd3", "m3" ), 0.764555},
                    {("m3", "cm3" ), 1000000},
                    {("m3", "mm3" ), 1000000000},
                    {("m3", "dm3" ), 1000},
                    {("m3", "ft3" ), 35.3147},
                    {("m3", "in3" ), 61023.7},
                    {("m3", "yd3" ), 1.30795},
                    {("L","m3"),0.001 },
                    {("m3","L"),1000 },

                    {("ton", "kg" ), 1000},
                    {("g", "kg" ), 0.001},
                    {("kg", "ton" ), 0.001},
                    {("kg", "g" ), 1000},
                    {("ton", "lb" ), 2000},
                    {("lb", "ton" ), 0.0005},
                    {("lb", "kg" ), 0.453592},
                    {("kg", "lb" ), 2.20462},

                    {("m", "cm" ), 100},
                    {("cm", "m" ), 0.01},
                    {("m", "mm" ), 1000},
                    {("mm", "m" ), 0.001},
                    {("m", "km" ), 0.001},
                    {("km", "m" ), 1000},
                    {("m", "ft" ), 3.28084},
                    {("ft", "m" ), 0.3048},
                    {("m", "in" ), 39.3701},
                    {("in", "m" ), 0.0254},
                    {("m", "yd" ), 1.09361},
                    {("yd", "m" ), 0.9144},
                    {("cm", "mm" ), 10},
                    {("mm", "cm" ), 0.1},
                    {("cm", "km" ), 0.00001},
                    {("km", "cm" ), 100000},
                    {("cm", "ft" ), 0.0328084},
                    {("ft", "cm" ), 30.48},
                    {("cm", "in" ), 0.393701},
                    {("in", "cm" ), 2.54},
                    {("cm", "yd" ), 0.0109361},
                    {("yd", "cm" ), 91.44},
                    {("mm", "km" ), 0.000001},
                    {("km", "mm" ), 1000000},
                    {("mm", "ft" ), 0.00328084},
                    {("ft", "mm" ), 304.8},
                    {("mm", "in" ), 0.0393701},
                    {("in", "mm" ), 25.4},
                    {("mm", "yd" ), 0.00109361},
                    {("yd", "mm" ), 914.4},
                    {("km", "ft" ),3280.84},

                    {("m2", "mm2" ),1000000},
                    {("mm2", "m2" ),0.000001},
                    {("m2", "cm2" ),10000},
                    {("cm2", "m2" ),0.0001},
                    {("m2", "km2" ),0.000001},
                    {("km2", "m2" ),1000000},
                    {("m2", "ft2" ),10.7639},
                    {("ft2", "m2" ),0.092903},
                    {("m2", "in2" ),1550.0031},
                    {("in2", "m2" ),0.00064516},
                    {("m2", "yd2" ),1.19599},
                    {("yd2", "m2" ),0.836127},
                    {("mm2", "cm2" ),0.01},
                    {("cm2", "mm2" ),100},
                    {("mm2", "km2" ),0.000000001},
                    {("km2", "mm2" ),1000000000},
                    {("mm2", "ft2" ),0.0000107639},
                    {("ft2", "mm2" ),92903.04},
                    {("mm2", "in2" ),0.0015500031},
                    {("in2", "mm2" ),645.16},
                    {("mm2", "yd2" ),0.00000119599},
                    {("yd2", "mm2" ),836127},
                    {("cm2", "km2" ),0.0000000001},
                    {("km2", "cm2" ),10000000000},
                    {("cm2", "ft2" ),0.00107639},
                    {("ft2", "cm2" ),929.0304},
                    {("cm2", "in2" ),0.15500031},
                    {("in2", "cm2" ),6.4516},
                    {("cm2", "yd2" ),0.000119599},
                    {("yd2", "cm2" ),8361.27},
                    {("km2", "ft2" ),10763910.4167},
                    {("ft2", "km2" ),0.000000092903},
                    {("km2", "in2" ),1.55e+9},
                    {("dm2","m2"), 0.01},
                    {("m2","dm2"), 100},
                };
        */
        static readonly Dictionary<string, double> ConversionVolume = new Dictionary<string, double>
        {
            { "m3",1 },
            { "cm3",0.000001 },
            { "mm3",0.000000001 },
            { "dm3",0.001 },
            { "ft3",0.0283168 },
            { "in3",0.0000163871 },
            { "yd3",0.764555 },
            { "L",0.001 },
        };

        static readonly Dictionary<string, double> ConversionArea = new Dictionary<string, double>
        {
            { "m2",1 },
            { "cm2",0.0001 },
            { "mm2",0.000001 },
            { "dm2",0.01 },
            {"km2",1000000},
            { "ft2",0.092903 },
            { "in2",0.00064516 },
            { "yd2",0.836127 },
        };

        static readonly Dictionary<string, double> ConversionWeight = new Dictionary<string, double>
        {
            { "ton",1000 },
            { "kg",1 },
            { "g",0.001 },
            { "lb",0.453592 },
        };

        static readonly Dictionary<string, double> ConversionLength = new Dictionary<string, double>
        {
            { "m",1 },
            { "cm",0.01 },
            { "mm",0.001 },
            { "dm",0.1 },
            { "km",1000 },
            { "ft",0.3048 },
            { "in",0.0254 },
            { "yd",0.9144 },
        };
    }
}
