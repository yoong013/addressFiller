// Models/PostcodeMap.cs
using System.Collections.Generic;  

namespace AddressFilter.Models
{
    public class State
    {
        public string name { get; set; } = string.Empty;
        
        public List<CityInfo> city { get; set; } = new List<CityInfo>();
    }

    public class CityInfo
    {
        public string name { get; set; } = string.Empty;
        
        public List<string> postcode { get; set; } = new List<string>();
    }
}
