using Newtonsoft.Json;
using System.Collections.Generic;

namespace SignupGeniusPDF
{
    public class SignUp
    {
        public string Title { get; set; }
        public string SignupID { get; set; }
    }
    public class FullSignUpReturn
    {
        [JsonProperty("message")]
        public List<object> Message { get; set; }
        [JsonProperty("success")]
        public bool Success { get; set; }
        [JsonProperty("data")]
        public SignUp[] Data { get; set; }
    }
}
