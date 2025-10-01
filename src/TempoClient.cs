using System.Net.Http.Headers;
using System.Net.Http.Json;

namespace Tempo.Exporter
{
    internal class TempoClient
    {
        private const string API_URL = "https://api.tempo.io";
        private const string WORKLOG_ENDPOINT = "/4/worklogs/user/{0}?from={1}&to={2}&limit={3}";
        private const string API_DATE_FORMAT = "yyyy-MM-dd";

        private const int WORKLOG_LIMIT = 10_000;

        private string _accountId;
        private string _apiKey;

        public TempoClient(string accountId, string apiKey)
        {
            _accountId = accountId;
            _apiKey = apiKey;
        }

        public async Task<Dictionary<DateTime, IEnumerable<TimeRange>>> GetWorklogsAsync(DateTime from, DateTime until)
        {
            using var httpClient = CreateTempoHttpClient();

            var response = await httpClient.GetFromJsonAsync<TempoResponseModel>(string.Format(WORKLOG_ENDPOINT, _accountId, from.ToString(API_DATE_FORMAT), until.ToString(API_DATE_FORMAT), WORKLOG_LIMIT));

            if (response is null)
            {
                return [];
            }

            return response.Results
                .GroupBy(a => DateTime.Parse(a.StartDate))
                .OrderBy(a => a.Key)
                .ToDictionary
                (
                    k => k.Key,
                    v => v.Select(k => new TimeRange(TimeSpan.Parse(k.StartTime), TimeSpan.Parse(k.StartTime).Add(TimeSpan.FromSeconds(k.TimeSpentSeconds)), k?.Issue?.Id ?? 0))
                );
        }

        private HttpClient CreateTempoHttpClient()
        {
            var httpClient = new HttpClient
            {
                BaseAddress = new Uri(API_URL)
            };

            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

            return httpClient;
        }
    }
}
