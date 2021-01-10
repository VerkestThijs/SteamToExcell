using System;
using System.Collections.Generic;
using System.Text;

namespace SteamToExcell
{
  
    
        public class Game
        {
            public int appid { get; set; }
            public string name { get; set; }
            public int playtime_forever { get; set; }
            public int playtime_windows_forever { get; set; }
            public int playtime_mac_forever { get; set; }
            public int playtime_linux_forever { get; set; }
        }

        public class Response
        {
            public int game_count { get; set; }
            public List<Game> games { get; set; }
        }

        public class Root
        {
            public Response response { get; set; }
        }
    
}
