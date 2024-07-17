using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;


namespace HelloWorld
{
    class OpenRouteServiceClient
    {
        private static readonly HttpClient client = new HttpClient();
        private readonly string apiKey;

        public OpenRouteServiceClient(string apiKey)
        {
            this.apiKey = apiKey;
        }

        public async Task<(double lat, double lon)> GeocodeAddressAsync(string address)
        {
            var requestUrl = $"https://api.openrouteservice.org/geocode/search?api_key={apiKey}&text={Uri.EscapeDataString(address)}";

            var response = await client.GetAsync(requestUrl);
            var message= response.EnsureSuccessStatusCode();

            if(message.StatusCode != System.Net.HttpStatusCode.OK)
            {
                return (0,0);
            }

            var responseContent = await response.Content.ReadAsStringAsync();
            var json = JObject.Parse(responseContent);

            double lat = json["features"][0]["geometry"]["coordinates"][1].Value<double>();
            double lon = json["features"][0]["geometry"]["coordinates"][0].Value<double>();
            return (lat, lon);
        }

        public async Task<double> GetDistanceAsync(double startLat, double startLon, double endLat, double endLon)
        {
            var requestUrl = $"https://api.openrouteservice.org/v2/directions/driving-car?api_key={apiKey}&start={startLon},{startLat}&end={endLon},{endLat}";

            var response = await client.GetAsync(requestUrl);

            response.EnsureSuccessStatusCode();

            var responseContent = await response.Content.ReadAsStringAsync();
            var json = JObject.Parse(responseContent);

            // Convert JObject to JSON string
/*            string jsonString = json.ToString();

            // Specify the file path where you want to save the JSON file
            string filePath = "example.json";

            // Write the JSON string to the file
            File.WriteAllText(filePath, jsonString);*/

            double distance = json["features"][0]["properties"]["summary"]["distance"].Value<double>();

            return distance;
        }

    }

    class MaritimeTrafficClient
    {
        private static readonly string apiKey = "";

        private static async Task<string> GetMaritimeDistance(string portStart, string portEnd)
        {
            using (HttpClient client = new HttpClient())
            {
                string url = $"https://api.marinetraffic.com/api/portdistance/{apiKey}/start/{portStart}/end/{portEnd}";
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                return responseBody;
            }
        }

        public static async Task<double> CalculateMaritimeDistance(string portStart,string portEnd)
        {          
            string distanceData = await GetMaritimeDistance(portStart, portEnd);

            JObject json = JObject.Parse(distanceData);

            double distance = json["distance"].Value<double>();

            return distance;
        }
    }
}

