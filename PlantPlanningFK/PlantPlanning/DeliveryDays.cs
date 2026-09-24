using Newtonsoft.Json; // Если используете Newtonsoft, или System.Text.Json.Serialization для встроенного

namespace PlantPlanningFKBis.Context
{
    public class DeliveryDayItem
    {
        [JsonProperty("Article")]
        public long Article { get; set; }

        [JsonProperty("DaysForDelivery")]
        public int? DaysForDelivery { get; set; }
    }
}