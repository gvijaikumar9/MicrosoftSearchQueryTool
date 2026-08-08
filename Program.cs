using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Azure.Core;
using Azure.Identity;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();
var cfg = app.Configuration;

// Default to the well-known "Microsoft Graph Command Line Tools" public client id,
// so there is no app registration to create. You just sign in and consent, the
// same way Connect-MgGraph works. Override in appsettings.json to use your own app.
string clientId  = cfg["ClientId"] ?? "14d82eec-204b-4c2f-b7e8-296a70dab67e";
string tenantId  = cfg["TenantId"] ?? "organizations";
string listenUrl = cfg["Url"]      ?? "http://localhost:5089";

// This build's version (from the .csproj <Version>), used by the update check.
string appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";

string[] scopes =
{
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/ExternalItem.Read.All",
    "https://graph.microsoft.com/ExternalConnection.Read.All"
};

var jsonOpts = new JsonSerializerOptions
{
    PropertyNameCaseInsensitive = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    WriteIndented = true
};

// Interactive browser sign-in, triggered explicitly from the page's Sign in button.
var credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
{
    ClientId = clientId,
    TenantId = tenantId,
    RedirectUri = new Uri("http://localhost"),
    TokenCachePersistenceOptions = new TokenCachePersistenceOptions { Name = "fivenumber-search-query-tool" }
});
var tokenCtx = new TokenRequestContext(scopes);
AuthenticationRecord? authRecord = null;

var http = new HttpClient();
// Returns null until someone signs in, so endpoints can answer 401 rather than
// silently triggering a popup.
async Task<string?> TokenAsync() =>
    authRecord is null ? null : (await credential.GetTokenAsync(tokenCtx, default)).Token;

// Any unhandled exception (auth cancelled, network down, Graph unreachable) comes back
// as clean JSON the UI can display, never a raw 500 page.
app.UseExceptionHandler(errApp => errApp.Run(async context =>
{
    var ex = context.Features.Get<Microsoft.AspNetCore.Diagnostics.IExceptionHandlerPathFeature>()?.Error;
    context.Response.StatusCode  = StatusCodes.Status500InternalServerError;
    context.Response.ContentType = "application/json";
    await context.Response.WriteAsync(JsonSerializer.Serialize(new { error = ex?.Message ?? "Unexpected server error." }));
}));

app.UseDefaultFiles();
app.UseStaticFiles();

// --- sign-in ------------------------------------------------------------------
// Current sign-in state (no prompt).
app.MapGet("/api/me", () => Results.Json(new { signedIn = authRecord is not null, user = authRecord?.Username }));

// Explicit sign in: opens the browser for interactive sign-in and consent.
app.MapPost("/api/signin", async () =>
{
    authRecord = await credential.AuthenticateAsync(tokenCtx);
    return Results.Json(new { signedIn = true, user = authRecord.Username });
});

// Sign out: drop the in-memory session (the on-disk token cache stays, so the next
// sign in is silent).
app.MapPost("/api/signout", () =>
{
    authRecord = null;
    return Results.Json(new { signedIn = false });
});

// Sign in as a different user: a fresh credential with no persisted cache, so the
// account picker always shows instead of silently reusing the cached account.
app.MapPost("/api/switch", async () =>
{
    credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
    {
        ClientId    = clientId,
        TenantId    = tenantId,
        RedirectUri = new Uri("http://localhost")
    });
    authRecord = await credential.AuthenticateAsync(tokenCtx);
    return Results.Json(new { signedIn = true, user = authRecord.Username });
});

// POST /api/search -> Graph POST /search/query. Returns the request we sent plus the
// raw Graph response, so the UI can show both the Request and Response tabs verbatim.
app.MapPost("/api/search", async (HttpContext ctx) =>
{
    var input = await JsonSerializer.DeserializeAsync<SearchInput>(ctx.Request.Body, jsonOpts)
                ?? new SearchInput();

    var token = await TokenAsync();
    if (token is null)
    {
        ctx.Response.StatusCode  = StatusCodes.Status401Unauthorized;
        ctx.Response.ContentType = "application/json";
        await ctx.Response.WriteAsync("{\"error\":\"not_signed_in\"}");
        return;
    }

    var req = new Dictionary<string, object?>
    {
        ["entityTypes"] = input.EntityTypes,
        ["query"]       = new { queryString = input.QueryString },
        ["from"]        = input.From,
        ["size"]        = input.Size
    };
    if (input.Fields is { Length: > 0 })         req["fields"]         = input.Fields;
    if (input.ContentSources is { Length: > 0 }) req["contentSources"] = input.ContentSources;
    if (input.Sort is { Length: > 0 })
        req["sortProperties"] = input.Sort.Select(s => new { name = s.Name, isDescending = s.IsDescending }).ToArray();
    if (input.Aggregations is { Length: > 0 })
        req["aggregations"] = input.Aggregations.Select(f => new
        {
            field = f,
            size = 10,
            bucketDefinition = new { sortBy = "count", isDescending = true, minimumCount = 0 }
        }).ToArray();
    if (input.AggregationFilters is { Length: > 0 })
        req["aggregationFilters"] = input.AggregationFilters;
    if (input.Spelling)
        req["queryAlterationOptions"] = new { enableSuggestion = true, enableModification = false };

    var payload = new { requests = new[] { req } };

    using var msg = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/search/query")
    {
        Content = new StringContent(JsonSerializer.Serialize(payload, jsonOpts), Encoding.UTF8, "application/json")
    };
    msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

    var resp = await http.SendAsync(msg);
    var raw  = await resp.Content.ReadAsStringAsync();

    ctx.Response.StatusCode  = (int)resp.StatusCode;
    ctx.Response.ContentType = "application/json";
    var envelope = new
    {
        ok       = resp.IsSuccessStatusCode,
        request  = payload,
        response = SafeParse(raw)
    };
    await ctx.Response.WriteAsync(JsonSerializer.Serialize(envelope, jsonOpts));
});

// GET /api/connections -> list the tenant's Copilot connectors so the UI can offer them
// in a dropdown instead of making you paste an id.
app.MapGet("/api/connections", async () =>
{
    var token = await TokenAsync();
    if (token is null) return Results.Json(new { error = "not_signed_in" }, statusCode: 401);
    using var msg = new HttpRequestMessage(HttpMethod.Get,
        "https://graph.microsoft.com/v1.0/external/connections?$select=id,name,state");
    msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    var resp = await http.SendAsync(msg);
    return Results.Content(await resp.Content.ReadAsStringAsync(), "application/json", null, (int)resp.StatusCode);
});

// GET /api/schema?connectionId=... -> the connection's property schema, so the UI can
// offer real, retrievable field names instead of guesses.
app.MapGet("/api/schema", async (string connectionId) =>
{
    var token = await TokenAsync();
    if (token is null) return Results.Json(new { error = "not_signed_in" }, statusCode: 401);
    using var msg = new HttpRequestMessage(HttpMethod.Get,
        $"https://graph.microsoft.com/v1.0/external/connections/{Uri.EscapeDataString(connectionId)}/schema");
    msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    var resp = await http.SendAsync(msg);
    return Results.Content(await resp.Content.ReadAsStringAsync(), "application/json", null, (int)resp.StatusCode);
});

// GET /api/root -> the tenant's SharePoint root URL, used to build the "verify in
// SharePoint enterprise search" links.
app.MapGet("/api/root", async () =>
{
    var token = await TokenAsync();
    if (token is null) return Results.Json(new { error = "not_signed_in" }, statusCode: 401);
    using var msg = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/sites/root?$select=webUrl");
    msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    var resp = await http.SendAsync(msg);
    return Results.Content(await resp.Content.ReadAsStringAsync(), "application/json", null, (int)resp.StatusCode);
});

// GET /api/sites?search=... -> search SharePoint sites for the "specific site" picker.
app.MapGet("/api/sites", async (string? search) =>
{
    var token = await TokenAsync();
    if (token is null) return Results.Json(new { error = "not_signed_in" }, statusCode: 401);
    var q = Uri.EscapeDataString(search ?? "");
    using var msg = new HttpRequestMessage(HttpMethod.Get,
        $"https://graph.microsoft.com/v1.0/sites?search={q}&$select=id,name,displayName,webUrl&$top=15");
    msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    var resp = await http.SendAsync(msg);
    return Results.Content(await resp.Content.ReadAsStringAsync(), "application/json", null, (int)resp.StatusCode);
});

// GET /api/update-check -> compares this build's version to the latest GitHub release,
// so the UI can flag when a newer version is available (and link to it). Runs server-side
// to avoid browser CORS and to send the User-Agent header the GitHub API requires.
app.MapGet("/api/update-check", async () =>
{
    const string repo    = "gvijaikumar9/MicrosoftSearchQueryTool";
    const string repoUrl = "https://github.com/gvijaikumar9/MicrosoftSearchQueryTool";
    try
    {
        using var msg = new HttpRequestMessage(HttpMethod.Get,
            $"https://api.github.com/repos/{repo}/releases/latest");
        msg.Headers.UserAgent.ParseAdd($"MicrosoftSearchQueryTool/{appVersion}");
        msg.Headers.Accept.ParseAdd("application/vnd.github+json");
        // Bound the call so a slow/unreachable GitHub can never hang the request.
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        var resp = await http.SendAsync(msg, cts.Token);
        // No releases published yet (404) or rate-limited -> just report "no update".
        if (!resp.IsSuccessStatusCode)
            return Results.Json(new { current = appVersion, latest = (string?)null, updateAvailable = false, url = repoUrl });
        var doc  = JsonSerializer.Deserialize<JsonElement>(await resp.Content.ReadAsStringAsync());
        var tag  = doc.TryGetProperty("tag_name", out var t) ? t.GetString() : null;
        var html = doc.TryGetProperty("html_url", out var h) ? h.GetString() : repoUrl;
        var newer = tag is not null && CompareVersions(tag, appVersion) > 0;
        return Results.Json(new { current = appVersion, latest = tag, updateAvailable = newer, url = html ?? repoUrl });
    }
    catch (Exception ex)
    {
        return Results.Json(new { current = appVersion, latest = (string?)null, updateAvailable = false, url = repoUrl, error = ex.Message });
    }
});

app.Urls.Add(listenUrl);
try { Process.Start(new ProcessStartInfo(listenUrl) { UseShellExecute = true }); } catch { /* open manually */ }
Console.WriteLine($"Microsoft Search Query Tool running at {listenUrl}");
app.Run();

static object? SafeParse(string s)
{
    try { return JsonSerializer.Deserialize<JsonElement>(s); }
    catch { return s; }
}

// Compares two dotted version strings, ignoring any leading "v". Returns >0 when a is newer
// than b, 0 when equal, <0 when older. Non-numeric segments (e.g. a "-beta" suffix) count as 0.
static int CompareVersions(string a, string b)
{
    static int[] Parts(string s) => s.TrimStart('v', 'V').Split('.', '-')
        .Select(p => int.TryParse(p, out var n) ? n : 0).ToArray();
    var pa = Parts(a);
    var pb = Parts(b);
    for (int i = 0; i < Math.Max(pa.Length, pb.Length); i++)
    {
        int x = i < pa.Length ? pa[i] : 0, y = i < pb.Length ? pb[i] : 0;
        if (x != y) return x.CompareTo(y);
    }
    return 0;
}

record SearchInput
{
    public string       QueryString    { get; init; } = "";
    public string[]     EntityTypes    { get; init; } = { "driveItem" };
    public string[]?    ContentSources { get; init; }
    public string[]?    Fields         { get; init; }
    public SortField[]? Sort               { get; init; }
    public string[]?    Aggregations       { get; init; }
    public string[]?    AggregationFilters { get; init; }
    public bool         Spelling           { get; init; }
    public int          From               { get; init; }
    public int          Size               { get; init; } = 25;
}

record SortField
{
    public string Name         { get; init; } = "";
    public bool   IsDescending { get; init; }
}
