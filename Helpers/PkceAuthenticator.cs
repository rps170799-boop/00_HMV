using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace HMVTools
{
    /// <summary>
    /// Holds the result of a successful PKCE authentication.
    /// Includes both access token (1-hour) and refresh token (15-day rolling).
    /// We don't persist the refresh token in Phase 2 — that comes in Phase 3.
    /// </summary>
    public class PkceTokenBundle
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }
    }

    /// <summary>
    /// Handles 3-Legged OAuth with PKCE for desktop apps.
    /// Opens default browser, listens on http://localhost:49152, exchanges code for tokens.
    /// No client_secret required — that's the whole point of PKCE.
    /// </summary>
    public class PkceAuthenticator
    {
        // ─── 🔑 PASTE YOUR NEW DESKTOP CLIENT ID HERE ────────────────
        // This is the Client ID from the new "HMV Batch Linker - Desktop" app
        // you created in the APS portal in Phase 1. NOT the old S2S one.
        private const string DesktopClientId = "qGwFrp16izSqNjwQVVyC2mwfaTwoQw1b72IUOQL1flprA2rn";

        // ─── Constants matching the APS portal config ─────────────────
        private const string CallbackUrl = "http://localhost:49152/api/auth/callback";
        private const string ListenerPrefix = "http://localhost:49152/";
        private const string AuthorizeEndpoint = "https://developer.api.autodesk.com/authentication/v2/authorize";
        private const string TokenEndpoint = "https://developer.api.autodesk.com/authentication/v2/token";
        private const string Scopes = "data:read account:read";
        private const int TimeoutMinutes = 5;

        //Token persistence paths -------------

        private static readonly string TokenFolder = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HMVTools");

        private static readonly string TokenFilePath = System.IO.Path.Combine(TokenFolder, "token.dat");



        /// <summary>
        /// Runs the full PKCE flow. Blocks until user completes login (or timeout).
        /// Returns the token bundle including refresh token for later persistence.
        /// </summary>
        public async Task<PkceTokenBundle> AuthenticateInteractiveAsync()
        {
            // 1. Generate PKCE pair (verifier stays in memory, challenge goes to Autodesk)
            string codeVerifier = GenerateCodeVerifier();
            string codeChallenge = GenerateCodeChallenge(codeVerifier);

            // 2. Generate state for CSRF protection (we'll verify it on callback)
            string state = GenerateRandomString(32);

            // 3. Build the authorize URL the user's browser will visit
            string authUrl = BuildAuthorizeUrl(codeChallenge, state);

            // 4. Start the local listener BEFORE opening the browser
            //    (otherwise there's a race where the redirect arrives before we're ready)
            using (HttpListener listener = new HttpListener())
            {
                listener.Prefixes.Add(ListenerPrefix);

                try
                {
                    listener.Start();
                }
                catch (HttpListenerException ex)
                {
                    throw new Exception(
                        $"Could not start local listener on port 49152. " +
                        $"Another app may be using this port, or Windows blocked it. " +
                        $"Original error: {ex.Message}");
                }

                // 5. Open the user's default browser to the Autodesk login page
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = authUrl,
                        UseShellExecute = true   // critical for .NET Framework — uses default browser
                    });
                }
                catch (Exception ex)
                {
                    listener.Stop();
                    throw new Exception($"Could not open default browser: {ex.Message}");
                }

                // 6. Wait for the callback (with timeout)
                string authorizationCode = await WaitForCallbackAsync(listener, state);

                // 7. Exchange the code for tokens
                PkceTokenBundle tokens = await ExchangeCodeForTokensAsync(authorizationCode, codeVerifier);

                // Save the refresh token for next time (no browser needed for ~15 days)
                SaveRefreshToken(tokens.RefreshToken);

                listener.Stop();
                return tokens;
            }






        }
        /// <summary>
        /// Tries to load a saved refresh token from disk and exchange it
        /// for a fresh access token — no browser needed.
        /// Returns null if no saved token, or if refresh fails (expired/revoked).
        /// </summary>
        public async Task<PkceTokenBundle> TryRefreshSilentlyAsync()
        {
            try
            {
                if (!System.IO.File.Exists(TokenFilePath))
                    return null;

                // Read and decrypt the refresh token
                byte[] encryptedBytes = System.IO.File.ReadAllBytes(TokenFilePath);
                byte[] decryptedBytes = System.Security.Cryptography.ProtectedData.Unprotect(
                    encryptedBytes,
                    null,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                string refreshToken = Encoding.UTF8.GetString(decryptedBytes);

                if (string.IsNullOrWhiteSpace(refreshToken))
                    return null;

                // Try exchanging the refresh token for a new access token
                using (HttpClient httpClient = new HttpClient())
                {
                    FormUrlEncodedContent body = new FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("grant_type", "refresh_token"),
                        new KeyValuePair<string, string>("client_id", DesktopClientId),
                        new KeyValuePair<string, string>("refresh_token", refreshToken)
                    });

                    HttpResponseMessage response = await httpClient.PostAsync(TokenEndpoint, body);
                    string responseJson = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        // Refresh token is expired or revoked — delete the stale file
                        try { System.IO.File.Delete(TokenFilePath); } catch { }
                        return null;
                    }

                    PkceTokenBundle tokens = JsonConvert.DeserializeObject<PkceTokenBundle>(responseJson);

                    // Save the NEW refresh token (Autodesk rolls it on each use)
                    SaveRefreshToken(tokens.RefreshToken);

                    return tokens;
                }
            }
            catch
            {
                // Any error (corrupt file, decryption failure, network) → fall back to interactive
                return null;
            }
        }

        /// <summary>
        /// Encrypts and saves the refresh token to disk.
        /// Uses Windows DPAPI (ProtectedData) — encrypted per-Windows-user,
        /// no admin rights needed, can't be read by other Windows accounts.
        /// </summary>
        public void SaveRefreshToken(string refreshToken)
        {
            try
            {
                if (!System.IO.Directory.Exists(TokenFolder))
                    System.IO.Directory.CreateDirectory(TokenFolder);

                byte[] plainBytes = Encoding.UTF8.GetBytes(refreshToken);
                byte[] encryptedBytes = System.Security.Cryptography.ProtectedData.Protect(
                    plainBytes,
                    null,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);

                System.IO.File.WriteAllBytes(TokenFilePath, encryptedBytes);
            }
            catch
            {
                // Non-fatal — if save fails, user just has to re-login next time
            }
        }

        /// <summary>
        /// Deletes the saved refresh token (for "sign out" functionality).
        /// </summary>
        public void ClearSavedToken()
        {
            try
            {
                if (System.IO.File.Exists(TokenFilePath))
                    System.IO.File.Delete(TokenFilePath);
            }
            catch { }
        }

        /// <summary>
        /// Blocks until the browser hits our /api/auth/callback URL.
        /// Validates state matches (CSRF check). Returns the authorization code.
        /// </summary>
        private async Task<string> WaitForCallbackAsync(HttpListener listener, string expectedState)
        {
            // Race: either the callback arrives, or the timeout fires
            Task<HttpListenerContext> contextTask = listener.GetContextAsync();
            Task timeoutTask = Task.Delay(TimeSpan.FromMinutes(TimeoutMinutes));

            Task completed = await Task.WhenAny(contextTask, timeoutTask);

            if (completed == timeoutTask)
            {
                throw new TimeoutException(
                    $"Login was not completed within {TimeoutMinutes} minutes. " +
                    $"Please try again.");
            }

            HttpListenerContext context = await contextTask;
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;

            // Parse query string from the redirect
            string code = request.QueryString["code"];
            string state = request.QueryString["state"];
            string error = request.QueryString["error"];
            string errorDescription = request.QueryString["error_description"];

            // Build the response shown in the browser tab
            string responseText;
            bool success = false;

            if (!string.IsNullOrEmpty(error))
            {
                responseText = $"Login failed: {error}. {errorDescription}. You can close this tab.";
            }
            else if (string.IsNullOrEmpty(code))
            {
                responseText = "Login failed: no authorization code received. You can close this tab.";
            }
            else if (state != expectedState)
            {
                responseText = "Login failed: state mismatch (possible CSRF attack). You can close this tab.";
            }
            else
            {
                responseText = "Login successful. You can close this tab and return to Revit.";
                success = true;
            }

            // Send the plain-text response to the browser
            byte[] buffer = Encoding.UTF8.GetBytes(responseText);
            response.ContentType = "text/plain; charset=utf-8";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Close();

            if (!success)
            {
                throw new Exception(responseText);
            }

            return code;
        }

        /// <summary>
        /// POST to Autodesk's token endpoint to trade the auth code for tokens.
        /// </summary>
        private async Task<PkceTokenBundle> ExchangeCodeForTokensAsync(string code, string codeVerifier)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                FormUrlEncodedContent body = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "authorization_code"),
                    new KeyValuePair<string, string>("client_id", DesktopClientId),
                    new KeyValuePair<string, string>("code", code),
                    new KeyValuePair<string, string>("code_verifier", codeVerifier),
                    new KeyValuePair<string, string>("redirect_uri", CallbackUrl)
                });

                HttpResponseMessage response = await httpClient.PostAsync(TokenEndpoint, body);
                string responseJson = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(
                        $"Token exchange failed: {(int)response.StatusCode} {response.StatusCode}\n" +
                        $"Response: {responseJson}");
                }

                return JsonConvert.DeserializeObject<PkceTokenBundle>(responseJson);
            }
        }

        // ─── PKCE helpers ────────────────────────────────────────────

        private string BuildAuthorizeUrl(string codeChallenge, string state)
        {
            // URL-encode each parameter to be safe
            string encodedClientId = Uri.EscapeDataString(DesktopClientId);
            string encodedRedirect = Uri.EscapeDataString(CallbackUrl);
            string encodedScopes = Uri.EscapeDataString(Scopes);

            return AuthorizeEndpoint
                + "?response_type=code"
                + "&client_id=" + encodedClientId
                + "&redirect_uri=" + encodedRedirect
                + "&scope=" + encodedScopes
                + "&code_challenge=" + codeChallenge
                + "&code_challenge_method=S256"
                + "&state=" + state;
        }

        /// <summary>
        /// PKCE code verifier: 43-128 chars, URL-safe random string.
        /// We use 64 chars from a 48-byte random source.
        /// </summary>
        private string GenerateCodeVerifier()
        {
            byte[] bytes = new byte[48];
            using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(bytes);
            }
            return Base64UrlEncode(bytes);
        }

        /// <summary>
        /// PKCE code challenge: SHA256(verifier), base64url-encoded.
        /// </summary>
        private string GenerateCodeChallenge(string verifier)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.ASCII.GetBytes(verifier));
                return Base64UrlEncode(hash);
            }
        }

        /// <summary>
        /// Random alphanumeric string for the OAuth 'state' parameter (CSRF protection).
        /// </summary>
        private string GenerateRandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            byte[] bytes = new byte[length];
            using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(bytes);
            }
            StringBuilder sb = new StringBuilder(length);
            foreach (byte b in bytes)
            {
                sb.Append(chars[b % chars.Length]);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Base64-URL encoding (no padding, +→-, /→_) per RFC 4648.
        /// Required by PKCE spec — standard base64 won't work.
        /// </summary>
        private string Base64UrlEncode(byte[] data)
        {
            return Convert.ToBase64String(data)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }
    }
}