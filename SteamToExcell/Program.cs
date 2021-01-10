using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.IO;
using EPPlusTest;
using OfficeOpenXml;
using SteamToExcell.Resources;

namespace SteamToExcell
{
    class Program
    {
        static void Main(string[] args)
        {

            string steamKey = key.steamkey; //steamkey is hidden in folder resources but in gitignore due to sensitive data
            Console.WriteLine("give your steam id");
            string steamId = Console.ReadLine();
            HttpClient client = CreateHTTPClient();
            List<Game> gamelist = GetGames(client, steamId, steamKey).Result;
            var workbook = InitSpreadSheet();
  
            var worksheet = workbook.Workbook.Worksheets[0];
            
            int counter = 1;
            Console.WriteLine("creating...");
            foreach (Game game in gamelist)
            {
                EnterDataInSpreadSheet(game,counter,workbook,worksheet);
                counter++;
            }

            Console.WriteLine("done");
            Console.ReadLine();
        }


        private static HttpClient CreateHTTPClient()
        {
           
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://api.steampowered.com/IPlayerService");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            
            return client;
        }

        private async static Task<List<Game>> GetGames( HttpClient client, string steamid,string steamKey)
        {
            List<Game> games = new List<Game>();
            string jsonstring = "";
            string pathString = "/IPlayerService/GetOwnedGames/v0001/?key=" + steamKey + "&steamid=" + steamid + "&format=json&include_appinfo=true";
       
            HttpResponseMessage response = await client.GetAsync(pathString);  // get the response of the get call
            if (response.IsSuccessStatusCode)
            {
                jsonstring = await response.Content.ReadAsStringAsync();
                var returnedGames = JsonConvert.DeserializeObject<Root>(jsonstring); //deserialize the json string and place it in the root class
                foreach (var game in returnedGames.response.games)
                {
                    games.Add(game);  // place every returned game in the list
                }
            }
            return games;
        }

        private static ExcelPackage InitSpreadSheet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;   //the library used eppplus expects you do define the license.
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var package = new ExcelPackage(new FileInfo(path+"/gameslist.xlsx"));
            if (package.File.Exists)
            {
                return package;
            }
            var ws = package.Workbook.Worksheets.Add("gamesSheet");
            int amountOfColums = 3;
            for (int i = 1; i <= amountOfColums; i++)
            {
                ws.Column(i).AutoFit();   //make sure all columns are fitted to fully show the data inside of the columns.
            }
            return package;
        }

        private static void EnterDataInSpreadSheet(Game game, int counter, ExcelPackage wb, ExcelWorksheet ws)
        {

            if (counter == 1)
            {
                ws.Cells["A" + counter].Value = "Games";
                ws.Cells["A" + counter].Style.Font.UnderLine = true;
                ws.Cells["A" + counter].Style.Font.Bold = true;

                ws.Cells["B" + counter].Value = "Platform";
                ws.Cells["B" + counter].Style.Font.UnderLine = true;
                ws.Cells["B" + counter].Style.Font.Bold = true;

                ws.Cells["C" + counter].Value = "Time Played";
                ws.Cells["C" + counter].Style.Font.UnderLine = true;
                ws.Cells["C" + counter].Style.Font.Bold = true;

            }
            else
            {
                ws.Cells["A" + counter].Value = game.name;
                ws.Cells["B" + counter].Value = "Steam";
                var hours = Math.Ceiling((Decimal) game.playtime_forever / 60); // the api returns playtime in minutes, this changes it to hours which is closer to values steam gives
                ws.Cells["C" + counter].Value = hours + " hours played";
            }

            wb.Save();

        }
    }
}
