using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Tempo.Exporter
{
    class TempoResponseModel
    {
        [JsonPropertyName("results")]
        public List<Worklog> Results { get; set; }
    }

    class Worklog
    {
        [JsonPropertyName("attributes")]
        public Attributes Attributes { get; set; }

        [JsonPropertyName("author")]
        public Author Author { get; set; }

        [JsonPropertyName("billableSeconds")]
        public long BillableSeconds { get; set; }

        [JsonPropertyName("createdAt")]
        public DateTime CreatedAt { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("issue")]
        public Issue Issue { get; set; }

        [JsonPropertyName("self")]
        public string Self { get; set; }

        [JsonPropertyName("startDate")]
        public string StartDate { get; set; }

        [JsonPropertyName("startDateTimeUtc")]
        public DateTime StartDateTimeUtc { get; set; }

        [JsonPropertyName("startTime")]
        public string StartTime { get; set; }

        [JsonPropertyName("tempoWorklogId")]
        public long TempoWorklogId { get; set; }

        [JsonPropertyName("timeSpentSeconds")]
        public long TimeSpentSeconds { get; set; }

        [JsonPropertyName("updatedAt")]
        public DateTime UpdatedAt { get; set; }
    }
    class Attributes
    {
        [JsonPropertyName("self")]
        public string Self { get; set; }

        [JsonPropertyName("values")]
        public List<AttributeValue> Values { get; set; }
    }

    class AttributeValue
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("value")]
        public string Value { get; set; }
    }

    class Author
    {
        [JsonPropertyName("accountId")]
        public string AccountId { get; set; }

        [JsonPropertyName("self")]
        public string Self { get; set; }
    }

    class Issue
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("self")]
        public string Self { get; set; }
    }
}
