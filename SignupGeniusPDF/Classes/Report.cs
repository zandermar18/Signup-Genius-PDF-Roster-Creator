using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace SignupGeniusPDF
{

    public class FullReportReturn
    {
        [JsonProperty("message")]
        public List<object> Message { get; set; }
        [JsonProperty("success")]
        public bool Success { get; set; }
        [JsonProperty("data")]
        public Data DataSection  { get; set; }
    }
    public class ReportSlot
    {
        [JsonProperty("email")]
        public string Email { get; set; }
        [JsonProperty("firstname")]
        public string FirstName { get; set; }
        [JsonProperty("lastname")]
        public string LastName { get; set; }
        [JsonProperty("comment")]
        public string Comment { get; set; }
        [JsonProperty("customfields")]
        public List<CustomField> CustomQuestions { get; set; }
        [JsonIgnore]
        public DateTime StartTime { get; set; }
        [JsonProperty("startdatestring")]
        public string StartTimeString { get; set; }
        [JsonProperty("item")]
        public string SlotName { get; set; }
    }
    public class CustomField
    {
        [JsonProperty("customfieldid")]
        public string CustomFieldID { get; set; }
        [JsonProperty("value")]
        public string Value { get; set; }
    }
    public class Data
    {
        [JsonProperty("signup")]
        public ReportSlot[] Slots { get; set; }
    }
}
