using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace HelloWorld
{
    public class MapAPI
    {
/*        static async Task Main(string[] args)
        {
            string originAddress = "1600 Amphitheatre Parkway, Mountain View, CA";
            string destinationAddress = "1 Infinite Loop, Cupertino, CA";
            string apiKey = "YOUR_API_KEY"; // Replace with your OpenRouteService API key

            double distance = await GetDistanceAsync(originAddress, destinationAddress, apiKey);
            Console.WriteLine($"Distance: {distance} meters");
        }*/

        public static async Task<double> GetDistanceAsync(string origin, string destination, string apiKey= "5b3ce3597851110001cf6248ff1788b40b624a5c8ef99b7e4e089734")
        {
            string originCoords = await GetCoordinatesAsync(origin, apiKey);
            string destinationCoords = await GetCoordinatesAsync(destination, apiKey);

            using (HttpClient client = new HttpClient())
            {
                string requestUri = $"https://api.openrouteservice.org/v2/directions/driving-car?api_key={apiKey}&start={originCoords}&end={destinationCoords}";

                HttpResponseMessage response = await client.GetAsync(requestUri);
                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception("Error calling the OpenRouteService API for directions");
                }

                string responseContent = await response.Content.ReadAsStringAsync();
                JObject json = JObject.Parse(responseContent);

                double distanceInMeters = json["features"][0]["properties"]["segments"][0]["distance"].ToObject<double>();
                return distanceInMeters;
            }
        }

        static async Task<string> GetCoordinatesAsync(string address, string apiKey)
        {
            using (HttpClient client = new HttpClient())
            {
                string requestUri = $"https://api.openrouteservice.org/geocode/search?api_key={apiKey}&text={Uri.EscapeDataString(address)}";

                HttpResponseMessage response = await client.GetAsync(requestUri);
                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception("Error calling the OpenRouteService API for geocoding");
                }

                string responseContent = await response.Content.ReadAsStringAsync();
                JObject json = JObject.Parse(responseContent);

                double lon = json["features"][0]["geometry"]["coordinates"][0].ToObject<double>();
                double lat = json["features"][0]["geometry"]["coordinates"][1].ToObject<double>();
                return $"{lon},{lat}";
            }
        }
    }


}
