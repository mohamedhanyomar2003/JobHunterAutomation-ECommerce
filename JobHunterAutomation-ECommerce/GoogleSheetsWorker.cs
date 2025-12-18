using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using RestSharp;
using Microsoft.Extensions.Configuration;

public class GoogleSheetsWorker : BackgroundService
{
    private readonly ILogger<GoogleSheetsWorker> _logger;
    private readonly IConfiguration _configuration;

    public GoogleSheetsWorker(ILogger<GoogleSheetsWorker> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("🚀 Job Hunter Service Started...");

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _logger.LogInformation("👀 Checking Google Sheets at: {time}", DateTimeOffset.Now);
                ProcessSheets();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error occurred while processing sheets.");
            }

            await Task.Delay(60000, stoppingToken);
        }
    }

    private void ProcessSheets()
    {
        var spreadsheetId = _configuration["GoogleSheetsSettings:SpreadsheetId"];
        var sheetName = _configuration["GoogleSheetsSettings:SheetName"];

        var credentialPath = Path.Combine(AppContext.BaseDirectory, "credentials.json");

        if (!File.Exists(credentialPath))
        {
            _logger.LogWarning("⚠️ credentials.json not found!");
            return;
        }

        GoogleCredential credential;
        using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
        }

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "JobHunter",
        });

        var range = $"{sheetName}!A2:I";
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        var response = request.Execute();
        var values = response.Values;

        if (values != null && values.Count > 0)
        {
            for (int i = 0; i < values.Count; i++)
            {
                var row = values[i];
                if (row.Count < 3) continue;

                string fName = GetVal(row, 0);
                string lName = GetVal(row, 1);
                string email = GetVal(row, 2);
                string job = GetVal(row, 3);
                string comp = GetVal(row, 4);
                string link = GetVal(row, 5);
                string ind = GetVal(row, 6);
                string size = GetVal(row, 7);
                string status = GetVal(row, 8);

                if (!string.IsNullOrEmpty(email) &&
                    !string.IsNullOrEmpty(comp) &&
                    !string.IsNullOrEmpty(job) &&
                    status != "Sent")
                {
                    _logger.LogInformation($"💡 Found New Candidate: {fName} @ {comp}");

                    bool success = SendToHubSpot(fName, lName, email, job, comp, link, ind, size);

                    if (success)
                    {
                        UpdateRowStatus(service, spreadsheetId, sheetName, i + 2, "Sent");
                        _logger.LogInformation("✅ Pushed to HubSpot successfully.");
                    }
                }
            }
        }
    }

    private bool SendToHubSpot(string fName, string lName, string email, string job, string comp, string link, string ind, string size)
    {
        var hubSpotToken = _configuration["HubSpotSettings:Token"];

        var client = new RestClient("https://api.hubapi.com/crm/v3/objects/contacts");
        var request = new RestRequest("", Method.Post);
        request.AddHeader("Authorization", $"Bearer {hubSpotToken}");
        request.AddHeader("Content-Type", "application/json");

        var body = new
        {
            properties = new
            {
                email = email,
                firstname = fName,
                lastname = lName,
                jobtitle = job,
                company = comp,
                hs_linkedin_url = link,
                industry = ind,
                company_size = size
            }
        };

        request.AddJsonBody(body);

        try
        {
            var response = client.Execute(request);
            if (response.IsSuccessful || response.Content.Contains("Contact already exists"))
            {
                return true;
            }

            _logger.LogError($"HubSpot Error: {response.Content}");
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Connection Error: {ex.Message}");
            return false;
        }
    }

    private string GetVal(IList<object> row, int index) => row.Count > index ? row[index].ToString().Trim() : "";

    private void UpdateRowStatus(SheetsService service, string spreadsheetId, string sheetName, int rowNumber, string status)
    {
        var range = $"{sheetName}!I{rowNumber}";
        var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { status } } };
        var update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
        update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        update.Execute();
    }
}