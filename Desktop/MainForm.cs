using Microsoft.Identity.Client;
using Microsoft.Web.WebView2.WinForms;
using System.Net.Http.Headers;
using System.Text.Json;

namespace TeamsReportExporterDesktop;

public sealed class MainForm : Form
{
    private static readonly string[] GraphScopes = ["User.Read", "User.Read.All", "Reports.Read.All"];
    private readonly WebView2 browser = new();
    private readonly HttpClient graph = new();

    public MainForm()
    {
        Text = "Teams Report Exporter";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(900, 620);
        ClientSize = new Size(1280, 820);
        Icon = SystemIcons.Application;
        browser.Dock = DockStyle.Fill;
        Controls.Add(browser);
        Shown += LoadApplication;
    }

    private async void LoadApplication(object? sender, EventArgs e)
    {
        try
        {
            var userData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TeamsReportExporter", "WebView2");
            var environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userData);
            await browser.EnsureCoreWebView2Async(environment);
            browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            browser.CoreWebView2.WebMessageReceived += HandleWebMessage;
            browser.CoreWebView2.Navigate(new Uri(Path.Combine(AppContext.BaseDirectory, "wwwroot", "index.html")).AbsoluteUri);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Microsoft Edge WebView2 is required to run this application.\n\n{ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private async void HandleWebMessage(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs args)
    {
        try
        {
            using var document = JsonDocument.Parse(args.WebMessageAsJson);
            var root = document.RootElement;
            if (root.GetProperty("type").GetString() == "signIn")
                await SignInAndLoadReports(root.GetProperty("tenantId").GetString()!, root.GetProperty("clientId").GetString()!);
        }
        catch (Exception ex)
        {
            SendMessage(new { type = "authError", message = ex.Message });
        }
    }

    private async Task SignInAndLoadReports(string tenantId, string clientId)
    {
        try
        {
            var app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithDefaultRedirectUri()
                .Build();
            var account = (await app.GetAccountsAsync()).FirstOrDefault();
            AuthenticationResult token;
            try
            {
                token = await app.AcquireTokenSilent(GraphScopes, account).ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                token = await app.AcquireTokenInteractive(GraphScopes)
                    .WithParentActivityOrWindow(Handle)
                    .WithPrompt(Prompt.SelectAccount)
                    .ExecuteAsync();
            }

            graph.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
            var rows = await GetTeamsUserActivity(token.AccessToken);
            SendMessage(new { type = "graphData", rows, accountName = token.Account?.Username ?? "Connected user", tenantName = tenantId });
        }
        catch (MsalException ex)
        {
            SendMessage(new { type = "authError", message = $"Entra ID sign-in failed ({ex.ErrorCode}). Check the tenant, client ID, redirect URI, and admin consent." });
        }
        catch (HttpRequestException ex)
        {
            SendMessage(new { type = "authError", message = $"Microsoft Graph could not be reached: {ex.Message}" });
        }
    }

    private async Task<List<ReportRow>> GetTeamsUserActivity(string accessToken)
    {
        using var response = await graph.GetAsync("https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D30')");
        response.EnsureSuccessStatusCode();
        var csv = await response.Content.ReadAsStringAsync();
        var lines = csv.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return [];
        var headers = ParseCsvLine(lines[0]);
        var index = headers.Select((name, index) => (name, index)).ToDictionary(x => x.name.Trim(), x => x.index, StringComparer.OrdinalIgnoreCase);
        var rows = new List<ReportRow>();
        foreach (var line in lines.Skip(1))
        {
            var values = ParseCsvLine(line);
            string Value(string name) => index.TryGetValue(name, out var i) && i < values.Count ? values[i] : "";
            rows.Add(new ReportRow(Value("User Principal Name"), Value("User Display Name"), "", Value("Last Activity Date"), "", Value("Is Licensed")));
        }
        return rows;
    }

    private static List<string> ParseCsvLine(string line)
    {
        var values = new List<string>();
        var current = new System.Text.StringBuilder();
        var quoted = false;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] == '"' && (i + 1 >= line.Length || line[i + 1] != '"')) quoted = !quoted;
            else if (line[i] == '"' && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
            else if (line[i] == ',' && !quoted) { values.Add(current.ToString()); current.Clear(); }
            else current.Append(line[i]);
        }
        values.Add(current.ToString());
        return values;
    }

    private void SendMessage(object message)
    {
        if (browser.CoreWebView2 is not null)
            browser.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(message, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase }));
    }

    private sealed record ReportRow(string UserPrincipalName, string DisplayName, string PhoneNumber, string VoiceEnabled, string Location, string License);
}
