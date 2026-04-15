using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace HMVTools
{
    // This tiny class holds the token data Autodesk sends back
    public class ApsTokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }
    }

    public class CloudRevitFile
    {
        public string Name { get; set; }
        public string Urn { get; set; } // The hidden Cloud ID
        public string Path { get; set; }
    }

    public class ApsManager
    {
        // 🛑 PASTE YOUR KEYS HERE 🛑
        private readonly string _clientId = "GdnqzjLBaQ6334P3VvpnrDiC5OiUB2pLquTcUjeu42rCY92U";
        private readonly string _clientSecret = "pe0dMLb0uRO7tCAX6nhbcXkG8QJ1SlVvwMX0r30dgtw4lBChz5aZTo6LWnNTkMjZ";

        private readonly HttpClient _httpClient;
        public string CurrentToken { get; private set; } = string.Empty;

        public ApsManager()
        {
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Trades your keys for a 1-hour data:read VIP pass.
        /// </summary>
        public async Task<bool> AuthenticateAsync()
        {
            try
            {
                string authUrl = "https://developer.api.autodesk.com/authentication/v2/token";

                var requestData = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                    new KeyValuePair<string, string>("scope", "data:read")
                });

                HttpResponseMessage response = await _httpClient.PostAsync(authUrl, requestData);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    ApsTokenResponse tokenData = JsonConvert.DeserializeObject<ApsTokenResponse>(jsonResponse);

                    CurrentToken = tokenData.AccessToken;
                    return true; // Success!
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    System.Windows.MessageBox.Show($"APS Auth Failed: {response.StatusCode}\n{error}", "HMV Tools");
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Web error: {ex.Message}", "HMV Tools");
                return false;
            }
        }

        /// <summary>
        /// DIAGNOSTIC: Tests what the 2-legged token can actually see.
        /// Returns a big string with the raw JSON of both calls.
        /// </summary>
        public async Task<string> RunDiagnosticAsync(string apsProjectId)
        {
            var report = new System.Text.StringBuilder();

            // Make sure auth header is set on the client
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + CurrentToken);

            // ─── TEST 1: List hubs ───────────────────────────────────────────
            report.AppendLine("=== TEST 1: GET /project/v1/hubs ===");
            try
            {
                string hubsUrl = "https://developer.api.autodesk.com/project/v1/hubs";
                HttpResponseMessage hubsResp = await _httpClient.GetAsync(hubsUrl);
                report.AppendLine($"Status: {(int)hubsResp.StatusCode} {hubsResp.StatusCode}");
                string hubsJson = await hubsResp.Content.ReadAsStringAsync();
                report.AppendLine("Body:");
                report.AppendLine(hubsJson);
                report.AppendLine();

                // Try to extract first hub ID for test 2
                string firstHubId = null;
                if (hubsResp.IsSuccessStatusCode)
                {
                    try
                    {
                        dynamic hubsData = JsonConvert.DeserializeObject(hubsJson);
                        if (hubsData.data != null && hubsData.data.Count > 0)
                        {
                            firstHubId = hubsData.data[0].id.ToString();
                            report.AppendLine($">>> Extracted first hub ID: {firstHubId}");
                            report.AppendLine();
                        }
                        else
                        {
                            report.AppendLine(">>> No hubs returned in data array.");
                            report.AppendLine();
                        }
                    }
                    catch (Exception parseEx)
                    {
                        report.AppendLine($">>> Parse error: {parseEx.Message}");
                    }
                }

                // ─── TEST 2: Top folders of YOUR project ─────────────────────
                report.AppendLine("=== TEST 2: GET /topFolders ===");
                if (string.IsNullOrEmpty(firstHubId))
                {
                    report.AppendLine("SKIPPED — no hub ID available from Test 1.");
                }
                else
                {
                    string topUrl = $"https://developer.api.autodesk.com/project/v1/hubs/{firstHubId}/projects/{apsProjectId}/topFolders";
                    report.AppendLine($"URL: {topUrl}");
                    HttpResponseMessage topResp = await _httpClient.GetAsync(topUrl);
                    report.AppendLine($"Status: {(int)topResp.StatusCode} {topResp.StatusCode}");
                    string topJson = await topResp.Content.ReadAsStringAsync();
                    report.AppendLine("Body:");
                    report.AppendLine(topJson);
                    report.AppendLine();

                    // Pretty-print folder names if successful
                    if (topResp.IsSuccessStatusCode)
                    {
                        try
                        {
                            dynamic topData = JsonConvert.DeserializeObject(topJson);
                            report.AppendLine(">>> Top folder names found:");
                            foreach (var folder in topData.data)
                            {
                                string fname = folder.attributes.name.ToString();
                                string fid = folder.id.ToString();
                                report.AppendLine($"   • {fname}  (id: {fid})");
                            }
                        }
                        catch (Exception parseEx)
                        {
                            report.AppendLine($">>> Parse error: {parseEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"EXCEPTION: {ex.Message}");
            }

            return report.ToString();
        }
        /// <summary>
        /// PHASE 2: 3-Legged PKCE authentication.
        /// Opens browser, user logs in, we get a USER token (not an app token).
        /// This token can see whatever the logged-in user can see in ACC —
        /// no per-project provisioning required.
        /// </summary>
        public async Task<bool> AuthenticateUserInteractiveAsync()
        {
            try
            {
                PkceAuthenticator pkce = new PkceAuthenticator();
                PkceTokenBundle tokens = await pkce.AuthenticateInteractiveAsync();

                CurrentToken = tokens.AccessToken;

                // We'll persist tokens.RefreshToken in Phase 3.
                // For now, just keep the access token in memory for this session.
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"PKCE login failed:\n\n{ex.Message}",
                    "HMV Tools - Login Error");
                return false;
            }
        }


        /// <summary>
        /// 1. Gets the first Hub (Company Account) your app has access to.
        /// </summary>
        public async Task<string> GetFirstHubIdAsync()
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + CurrentToken);

            string url = "https://developer.api.autodesk.com/project/v1/hubs";
            HttpResponseMessage response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            string json = await response.Content.ReadAsStringAsync();
            dynamic data = JsonConvert.DeserializeObject(json);

            return data.data[0].id.ToString();
        }

        /// <summary>
        /// 2. Finds the root "Project Files" folder for the active Revit project.
        /// </summary>
        public async Task<string> GetProjectFilesFolderIdAsync(string hubId, string projectId)
        {
            string url = $"https://developer.api.autodesk.com/project/v1/hubs/{hubId}/projects/{projectId}/topFolders";
            HttpResponseMessage response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            string json = await response.Content.ReadAsStringAsync();
            dynamic data = JsonConvert.DeserializeObject(json);

            foreach (var folder in data.data)
            {
                string folderName = folder.attributes.name.ToString();
                if (folderName == "Project Files")
                {
                    return folder.id.ToString();
                }
            }
            return null;
        }

        /// <summary>
        /// 3. Recursively scans folders politely to find all .rvt files.
        /// </summary>
        public async Task<List<CloudRevitFile>> FindRevitFilesPolitelyAsync(string projectId, string folderId, string currentPath = "")
        {
            var foundFiles = new List<CloudRevitFile>();

            string url = $"https://developer.api.autodesk.com/data/v1/projects/{projectId}/folders/{folderId}/contents";
            HttpResponseMessage response = await _httpClient.GetAsync(url);

            if (!response.IsSuccessStatusCode) return foundFiles;

            string json = await response.Content.ReadAsStringAsync();
            dynamic data = JsonConvert.DeserializeObject(json);

            foreach (var item in data.data)
            {
                string type = item.type.ToString();
                string name = item.attributes.displayName?.ToString();

                if (type == "folders")
                {
                    // It's a subfolder! Be polite to the server, then drill down.
                    await Task.Delay(200);

                    string subFolderId = item.id.ToString();
                    string newPath = string.IsNullOrEmpty(currentPath) ? name : $"{currentPath} / {name}";

                    var subFiles = await FindRevitFilesPolitelyAsync(projectId, subFolderId, newPath);
                    foundFiles.AddRange(subFiles);
                }
                else if (type == "items" && name != null && name.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                {
                    // We found a Revit file! Save its data.
                    foundFiles.Add(new CloudRevitFile
                    {
                        Name = name,
                        Urn = item.id.ToString(),
                        Path = currentPath
                    });
                }
            }

            return foundFiles;
        }

        /// <summary>
        /// PHASE 2 DIAGNOSTIC v3: with user token, decode JWT + fetch hubs fresh.
        /// Goal: figure out why /projects returns 403 for the user.
        /// </summary>
        /// 
        /// <summary>
        /// PHASE 2 DIAGNOSTIC v4: Try alternative endpoints that are friendlier
        /// to project-only members (no account-level membership required).
        /// </summary>
        public async Task<string> RunDiagnosticV4Async()
        {
            var report = new System.Text.StringBuilder();

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + CurrentToken);

            // ─── TEST 1: Legacy hubs endpoint (we know this returns 0) ────
            report.AppendLine("=== TEST 1 (baseline): GET /project/v1/hubs ===");
            await DumpEndpoint(report, "https://developer.api.autodesk.com/project/v1/hubs");

            // ─── TEST 2: ACC-specific user-info endpoint ──────────────────
            // Returns the current user's profile — proves the token is valid
            report.AppendLine("=== TEST 2: GET /userprofile/v1/users/@me ===");
            await DumpEndpoint(report, "https://api.userprofile.autodesk.com/userinfo");

            // ─── TEST 3: ACC project-based enumeration via user's projects ─
            // This newer endpoint enumerates projects the user is a MEMBER of,
            // bypassing the account-level requirement.
            report.AppendLine("=== TEST 3: GET /construction/admin/v1/projects/{projectId} ===");
            // We'll try the project ID we already know from Revit
            // (We can't enumerate all projects easily, but we CAN test if this single one is accessible)
            string knownProjectGuid = "af1e3723-2a56-4144-8f6e-e7813b3aff6d"; // bare GUID, no b. prefix
            await DumpEndpoint(report,
                $"https://developer.api.autodesk.com/construction/admin/v1/projects/{knownProjectGuid}");

            // ─── TEST 4: Direct project access via Data Management API ────
            // If this works, we can skip the hub/project enumeration entirely
            // and go straight to folder browsing using the project ID from Revit.
            report.AppendLine("=== TEST 4: GET /data/v1/projects/{projectId} ===");
            string apsProjectId = "b." + knownProjectGuid;
            await DumpEndpoint(report,
                $"https://developer.api.autodesk.com/data/v1/projects/{apsProjectId}");

            // ─── TEST 5: List root folders directly (skipping topFolders) ─
            // The /topFolders endpoint requires hub membership. But if we know
            // a folder URN, we can list contents directly. Let's see what the
            // project root looks like.
            report.AppendLine("=== TEST 5: Try project's roots via /hubs path with project (newer style) ===");
            // Newer ACC API exposes project info without hub navigation
            await DumpEndpoint(report,
                $"https://developer.api.autodesk.com/project/v1/projects/{apsProjectId}");

            return report.ToString();
        }

        /// <summary>
        /// Helper: GETs a URL with the current auth header, appends status + body to report.
        /// </summary>
        private async Task DumpEndpoint(System.Text.StringBuilder report, string url)
        {
            try
            {
                report.AppendLine($"URL: {url}");
                HttpResponseMessage resp = await _httpClient.GetAsync(url);
                report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                string body = await resp.Content.ReadAsStringAsync();
                if (body.Length > 1500)
                    body = body.Substring(0, 1500) + "... [truncated]";
                report.AppendLine("Body: " + body);
                report.AppendLine();
            }
            catch (Exception ex)
            {
                report.AppendLine($"EXCEPTION: {ex.Message}\n");
            }
        }
        public async Task<string> RunDiagnosticV3Async()
        {
            var report = new System.Text.StringBuilder();

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + CurrentToken);

            // ─── PART 1: Decode the JWT to see what's inside the token ───
            report.AppendLine("=== PART 1: Token introspection ===");
            try
            {
                string[] parts = CurrentToken.Split('.');
                if (parts.Length >= 2)
                {
                    string payload = parts[1];
                    // Base64URL → Base64 (add padding)
                    payload = payload.Replace('-', '+').Replace('_', '/');
                    switch (payload.Length % 4)
                    {
                        case 2: payload += "=="; break;
                        case 3: payload += "="; break;
                    }
                    byte[] payloadBytes = Convert.FromBase64String(payload);
                    string payloadJson = System.Text.Encoding.UTF8.GetString(payloadBytes);
                    report.AppendLine("JWT payload:");
                    report.AppendLine(payloadJson);
                }
                else
                {
                    report.AppendLine("Token doesn't look like a JWT (no dots).");
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"JWT decode failed: {ex.Message}");
            }
            report.AppendLine();

            // ─── PART 2: Fetch hubs FRESH with the user token ───
            report.AppendLine("=== PART 2: GET /project/v1/hubs (FRESH) ===");
            string detectedHubId = null;
            try
            {
                string url = "https://developer.api.autodesk.com/project/v1/hubs";
                HttpResponseMessage resp = await _httpClient.GetAsync(url);
                report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                string json = await resp.Content.ReadAsStringAsync();

                if (resp.IsSuccessStatusCode)
                {
                    dynamic data = JsonConvert.DeserializeObject(json);
                    int count = data.data != null ? (int)data.data.Count : 0;
                    report.AppendLine($">>> Hubs visible to user: {count}");
                    if (count > 0)
                    {
                        foreach (var hub in data.data)
                        {
                            string hname = hub.attributes?.name?.ToString() ?? "(no name)";
                            string hid = hub.id.ToString();
                            string hext = hub.attributes?.extension?.type?.ToString() ?? "(no type)";
                            report.AppendLine($"   • {hname}");
                            report.AppendLine($"     id:   {hid}");
                            report.AppendLine($"     type: {hext}");
                            if (detectedHubId == null) detectedHubId = hid;
                        }
                    }
                }
                else
                {
                    report.AppendLine("Body: " + json);
                }
                report.AppendLine();
            }
            catch (Exception ex) { report.AppendLine($"EXCEPTION: {ex.Message}\n"); }

            // ─── PART 3: List projects with the FRESH hub ID ───
            report.AppendLine("=== PART 3: GET /hubs/{freshId}/projects ===");
            if (string.IsNullOrEmpty(detectedHubId))
            {
                report.AppendLine("SKIPPED — no hub ID from Part 2.");
            }
            else
            {
                try
                {
                    string url = $"https://developer.api.autodesk.com/project/v1/hubs/{detectedHubId}/projects";
                    report.AppendLine($"URL: {url}");
                    HttpResponseMessage resp = await _httpClient.GetAsync(url);
                    report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                    string json = await resp.Content.ReadAsStringAsync();

                    if (resp.IsSuccessStatusCode)
                    {
                        dynamic data = JsonConvert.DeserializeObject(json);
                        int count = data.data != null ? (int)data.data.Count : 0;
                        report.AppendLine($">>> Projects visible to user in this hub: {count}");
                        if (count > 0)
                        {
                            report.AppendLine(">>> First few project names:");
                            int shown = 0;
                            foreach (var proj in data.data)
                            {
                                string pname = proj.attributes?.name?.ToString() ?? "(no name)";
                                string pid = proj.id.ToString();
                                report.AppendLine($"   • {pname}  ({pid})");
                                if (++shown >= 10) { report.AppendLine("   ... (truncated to first 10)"); break; }
                            }
                        }
                    }
                    else
                    {
                        report.AppendLine("Body: " + json);
                    }
                }
                catch (Exception ex) { report.AppendLine($"EXCEPTION: {ex.Message}"); }
            }
            report.AppendLine();

            return report.ToString();
        }
        /// <summary>
        /// DIAGNOSTIC v2: Tests project visibility from multiple angles.
        /// </summary>
        public async Task<string> RunDiagnosticV2Async(string apsProjectId, string hubId)
        {
            var report = new System.Text.StringBuilder();

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + CurrentToken);

            // Strip "b." prefix to get the bare GUID for some tests
            string bareGuid = apsProjectId.StartsWith("b.") ? apsProjectId.Substring(2) : apsProjectId;

            // ─── TEST A: List ALL projects in the hub ────────────────────────
            report.AppendLine("=== TEST A: List all projects in hub ===");
            try
            {
                string url = $"https://developer.api.autodesk.com/project/v1/hubs/{hubId}/projects";
                HttpResponseMessage resp = await _httpClient.GetAsync(url);
                report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                string json = await resp.Content.ReadAsStringAsync();

                if (resp.IsSuccessStatusCode)
                {
                    dynamic data = JsonConvert.DeserializeObject(json);
                    int count = data.data != null ? (int)data.data.Count : 0;
                    report.AppendLine($">>> Projects visible to app: {count}");
                    if (count > 0)
                    {
                        report.AppendLine(">>> Project names + IDs:");
                        foreach (var proj in data.data)
                        {
                            string pname = proj.attributes.name.ToString();
                            string pid = proj.id.ToString();
                            bool isMatch = pid.Contains(bareGuid);
                            string flag = isMatch ? "  <-- MATCHES YOUR ACTIVE PROJECT" : "";
                            report.AppendLine($"   • {pname}  (id: {pid}){flag}");
                        }
                    }
                }
                else
                {
                    report.AppendLine("Body: " + json);
                }
                report.AppendLine();
            }
            catch (Exception ex) { report.AppendLine($"EXCEPTION: {ex.Message}"); }

            // ─── TEST B: Try project detail with "b." prefix ─────────────────
            report.AppendLine("=== TEST B: Project detail with 'b.' prefix ===");
            try
            {
                string url = $"https://developer.api.autodesk.com/project/v1/hubs/{hubId}/projects/{apsProjectId}";
                HttpResponseMessage resp = await _httpClient.GetAsync(url);
                report.AppendLine($"URL: {url}");
                report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                string body = await resp.Content.ReadAsStringAsync();
                // Just show first 500 chars of body
                report.AppendLine("Body (truncated): " + (body.Length > 500 ? body.Substring(0, 500) + "..." : body));
                report.AppendLine();
            }
            catch (Exception ex) { report.AppendLine($"EXCEPTION: {ex.Message}"); }

            // ─── TEST C: Try project detail WITHOUT prefix ───────────────────
            report.AppendLine("=== TEST C: Project detail WITHOUT prefix (bare GUID) ===");
            try
            {
                string url = $"https://developer.api.autodesk.com/project/v1/hubs/{hubId}/projects/{bareGuid}";
                HttpResponseMessage resp = await _httpClient.GetAsync(url);
                report.AppendLine($"URL: {url}");
                report.AppendLine($"Status: {(int)resp.StatusCode} {resp.StatusCode}");
                string body = await resp.Content.ReadAsStringAsync();
                report.AppendLine("Body (truncated): " + (body.Length > 500 ? body.Substring(0, 500) + "..." : body));
                report.AppendLine();
            }
            catch (Exception ex) { report.AppendLine($"EXCEPTION: {ex.Message}"); }

            return report.ToString();
        }
    }
}