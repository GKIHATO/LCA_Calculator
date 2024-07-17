using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace HelloWorld
{

    public class DistanceCalculator
    {
        private readonly OpenRouteServiceClient orsClient;

        Dictionary<string, (double, double)> addressDictionary;

        Dictionary<(string, string), double> distanceDictionary;

        public DistanceCalculator(string apiKey)
        {
            orsClient = new OpenRouteServiceClient(apiKey);

            addressDictionary = new Dictionary<string, (double, double)>();

            distanceDictionary = new Dictionary<(string, string), double>();
        }

        public async Task<double> CalculateAndProceed(string startAddress, string endAddress)
        {

            (double lat,double lon) geoInfo_StartAddress;

            (double lat, double lon) geoInfo_EndAddress;
            // Geocode the addresses
            if (addressDictionary.Keys.Contains(startAddress))
            {
                geoInfo_StartAddress = addressDictionary[startAddress];
            }
            else
            {
                geoInfo_StartAddress = await orsClient.GeocodeAddressAsync(startAddress);

                addressDictionary.Add(startAddress,geoInfo_StartAddress);
            }

            if(addressDictionary.Keys.Contains(endAddress))
            {
                geoInfo_EndAddress = addressDictionary[endAddress];
            }
            else
            {
                geoInfo_EndAddress = await orsClient.GeocodeAddressAsync(endAddress);

                addressDictionary.Add(endAddress,geoInfo_EndAddress);
            }
            
            if(geoInfo_StartAddress == (0,0) || geoInfo_EndAddress == (0,0))
            {
                return 0.0;
            }

            // Call the async method and wait for the result

            double distance = 0;

            if(distanceDictionary.ContainsKey((startAddress,endAddress)))
            {
                distance = distanceDictionary[(startAddress,endAddress)];
            }
            else
            {
                distance = await orsClient.GetDistanceAsync(geoInfo_StartAddress.lat, geoInfo_StartAddress.lon, geoInfo_EndAddress.lat, geoInfo_EndAddress.lon);

                distanceDictionary.Add((startAddress,endAddress),distance);
            }

            return distance;
        }

        public double CalculateLandDistance(string startAddress, string endAddress)
        {
            double distance = 0;
            // Geocode the addresses
            Task.Run(async () =>
            {
                try
                {
                    distance = await CalculateAndProceed(startAddress, endAddress);
                    // Use the distance here
                }
                catch (Exception ex)
                {
                    distance = 0;
                    // Handle exceptions
                }
            }).Wait();

            if(distance == 0)
            {
                distance = 20000;
            }
            return distance;
        }

        public double CalculateMaritimeDistance(string portStart, string portEnd)
        {
            double distance = 0;
            // Geocode the addresses
            Task.Run(async () =>
            {
                try
                {
                    distance = await MaritimeTrafficClient.CalculateMaritimeDistance(portStart, portEnd);
                    // Use the distance here
                }
                catch (Exception ex)
                {
                    int i = 0;
                    // Handle exceptions
                }
            }).Wait();

            return distance;
        }

    }
    
}
