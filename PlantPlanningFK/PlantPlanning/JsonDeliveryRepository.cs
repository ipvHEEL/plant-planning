using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace PlantPlanningFKBis.Context
{
    public class JsonDeliveryRepository
    {
        private readonly string _filePath;

        public JsonDeliveryRepository()
        {
            _filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "delivery_days.json");
        }

        public List<DeliveryDayItem> GetAll()
        {
            if (!File.Exists(_filePath))
                return new List<DeliveryDayItem>();

            var json = File.ReadAllText(_filePath);
            return JsonConvert.DeserializeObject<List<DeliveryDayItem>>(json)
                   ?? new List<DeliveryDayItem>();
        }

        public void SaveAll(List<DeliveryDayItem> items)
        {
            var json = JsonConvert.SerializeObject(items, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(_filePath, json);
        }

        public DeliveryDayItem GetByArticle(long article)
        {
            return GetAll().FirstOrDefault(x => x.Article == article);
        }

        public void AddOrUpdate(DeliveryDayItem item)
        {
            var list = GetAll();
            var existing = list.FirstOrDefault(x => x.Article == item.Article);

            if (existing != null)
                list.Remove(existing);

            list.Add(item);
            SaveAll(list);
        }
    }
}