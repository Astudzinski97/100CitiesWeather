using OfficeOpenXml;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;

namespace _100CitiesWeather
{
    class Program
    {
        static void Main()
        {
            // Getting data from countriesnow api
            var client = new RestClient("https://countriesnow.space/api/v0.1/countries");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);

            // Cutting header from api response
            String longString = response.Content.Substring(62);

            // getting all cities names from first 50 countries in single String
            String cities = FormatData(longString);

            // choose randomly 100 cities from the list
            String[] cities100 = GetRandom100Cities(cities);

            // setting up arrays that will store weather data for our 100 cities (general weather, temperature, air pressure, humidity, wind speed and cloudiness)
            String[] weatherArray = new String[100];
            String[] tempArray = new String[100];
            String[] pressArray = new String[100];
            String[] humidArray = new String[100];
            String[] windArray = new String[100];
            String[] cloudArray = new String[100];


            // setting up RestClient to access openweathermap api
            var clientWeather = new RestClient();
            clientWeather.Timeout = -1;

            // looping through 100 random cities and downloading weather data from api, then assigning data to corresponding arrays
            for (int i = 0; i < cities100.Length; i++)
            {
                Console.WriteLine("App is getting data: " + (i+1) + "/100");
                String url = String.Format("https://api.openweathermap.org/data/2.5/weather?q={0}&appid=2188f233a042c1974a856519c162950a&units=metric", cities100[i]);
                clientWeather = new RestClient(url);
                response = clientWeather.Execute(request);
                // if city is not found in weather api it fills arrays with "city not found"
                if (!response.Content.Contains("city not found"))
                {
                    String[] weather = GetWeatherData(response.Content);
                    weatherArray[i] = weather[0];
                    tempArray[i] = weather[1];
                    pressArray[i] = weather[2];
                    humidArray[i] = weather[3];
                    windArray[i] = weather[4];
                    cloudArray[i] = weather[5];
                } else
                {
                    weatherArray[i] = "city not found";
                    tempArray[i] = "city not found";
                    pressArray[i] = "city not found";
                    humidArray[i] = "city not found";
                    windArray[i] = "city not found";
                    cloudArray[i] = "city not found";
                }
            }

            for (int j = 0; j < cities100.Length; j++)
            {
                Console.WriteLine((j + 1) + " " + cities100[j]);
                Console.WriteLine(" General: " + weatherArray[j]);
                Console.WriteLine(" Temperature: " + tempArray[j]);
                Console.WriteLine(" Air Pressure: " + pressArray[j]);
                Console.WriteLine(" Humidity: " + humidArray[j]);
                Console.WriteLine(" Wind Speed: " + windArray[j]);
                Console.WriteLine(" Cloudiness: " + cloudArray[j]);
            }

            Console.WriteLine("Making excel sheet");
            // saving all the data to excel sheet 
            // directory is hardset to current directory - 100CitiesWeather\100CitiesWeather\bin\Debug
            // if needed can change it to writing path from Console - would less errorproof and need some protection though
            // also need to set License for Epplus in order to work can do it in app.config but feels like currently it is more transparent
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(Directory.GetCurrentDirectory()+@"\myWorkbook.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.Add("Weather in 100 Cities " + DateTime.Now.ToString("HH:mm:ss"));
                sheet.Cells[1, 1].Value = "Number";
                sheet.Cells[1, 2].Value = "City Name";
                sheet.Cells[1, 3].Value = "General Weather";
                sheet.Cells[1, 4].Value = "Temperature";
                sheet.Cells[1, 5].Value = "Air Pressure";
                sheet.Cells[1, 6].Value = "Humidity";
                sheet.Cells[1, 7].Value = "Wind Speed";
                sheet.Cells[1, 8].Value = "Cloudiness";

                for(int k = 0; k < cities100.Length; k++)
                {
                    sheet.Cells[k + 2, 1].Value = k + 1;
                    sheet.Cells[k + 2, 2].Value = cities100[k];
                    sheet.Cells[k + 2, 3].Value = weatherArray[k];
                    sheet.Cells[k + 2, 4].Value = tempArray[k];
                    sheet.Cells[k + 2, 5].Value = pressArray[k];
                    sheet.Cells[k + 2, 6].Value = humidArray[k];
                    sheet.Cells[k + 2, 7].Value = windArray[k];
                    sheet.Cells[k + 2, 8].Value = cloudArray[k];
                }

                package.Save();
            }

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }

        static String FormatData(String longString)
        {
            // Formating response to get cities names only (only for first country)
            longString = longString.Remove(0, 33);
            int start = longString.IndexOf("[") + 1;
            int end = longString.IndexOf("]") - (longString.IndexOf("[") + 1);
            String cities = " ";
            cities += longString.Substring(start, end);

            // Adding cities from first 60 countries to have enough sample size
            for (int i = 0; i < 60; i++)
            {
                longString = longString.Remove(0, longString.IndexOf("]") + 1);
                start = longString.IndexOf("[") + 1;
                end = longString.IndexOf("]") - (longString.IndexOf("[") + 1);
                cities += "," + longString.Substring(start, end);
            }

            return cities;
        }

        static String[] GetRandom100Cities(String cities)
        {
            List<String> citiesList = new List<String>();

            // adding all cities names as individual strings into the list
            do {
                String temp = cities.Substring(cities.IndexOf("\"") + 1, cities.IndexOf(",") - 2);
                citiesList.Add(temp);
                cities = cities.Remove(0, cities.IndexOf(",") + 1);
            } while (cities.IndexOf(",") != -1);

            String[] citiesArray = new String[100];

            Random rand = new Random();
            int randPos;

            for ( int i = 0; i < citiesArray.Length; i++)
            {
                randPos = rand.Next(0, citiesList.Count);
                // ensure that city name doesnt have space in it - Weather API doesnt seem to like that
                while (citiesList[randPos].IndexOf(" ") != -1) { randPos = rand.Next(0, citiesList.Count); };
                citiesArray[i] = citiesList[randPos];
                citiesList.RemoveAt(randPos);
            }
            return citiesArray;
        }

        static String[] GetWeatherData(String response)
        {
            String[] weatherData = new String[6];

            // getting general weather data
            String general = response.Remove(0, response.IndexOf("main")+7);
            weatherData[0] = general.Substring(0, general.IndexOf("\""));

            // getting temperature
            general = response.Remove(0, response.IndexOf("temp")+6);
            weatherData[1] = general.Substring(0, general.IndexOf(",")) + "°C";

            // getting air pressure
            general = response.Remove(0, response.IndexOf("pressure") + 10);
            weatherData[2] = general.Substring(0, general.IndexOf(",")) +"hPa";

            //getting humidity
            general = response.Remove(0, response.IndexOf("humidity") + 10);
            weatherData[3] = general.Substring(0, Math.Min(general.IndexOf(","), general.IndexOf("}"))) + "%";

            // getting wind speed
            general = response.Remove(0, response.IndexOf("speed") + 7);
            weatherData[4] = general.Substring(0, general.IndexOf(",")) + "m/s";

            // getting cloudiness
            general = response.Remove(0, response.IndexOf("all") + 5);
            weatherData[5] = general.Substring(0, Math.Min(general.IndexOf(","), general.IndexOf("}"))) + "%";

            return weatherData;
        }
    }
}
