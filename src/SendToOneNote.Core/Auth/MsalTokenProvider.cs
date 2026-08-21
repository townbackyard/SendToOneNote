using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;

namespace SendToOneNote.Core.Auth;

public sealed class MsalTokenProvider : ITokenProvider
{
    public const string DefaultClientId = "00000000-0000-0000-0000-000000000000";
    private static readonly string[] Scopes = ["User.Read", "Notes.ReadWrite"];

    private readonly IPublicClientApplication _pca;
    private readonly SemaphoreSlim _init = new(1, 1);
    private bool _cacheAttached;
    private readonly string _cacheDir;

    public string? SignedInUser { get; private set; }

    public MsalTokenProvider(string cacheDir, string? clientIdOverride = null,
        IntPtr parentWindow = default)
    {
        _cacheDir = cacheDir;
        _pca = PublicClientApplicationBuilder
            .Create(string.IsNullOrWhiteSpace(clientIdOverride) ? DefaultClientId : clientIdOverride)
            .WithAuthority("https://login.microsoftonline.com/common")
            .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows))
            .WithParentActivityOrWindow(() => parentWindow)
            .Build();
    }

    public async Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default)
    {
        await EnsureCacheAsync();
        var accounts = await _pca.GetAccountsAsync();
        var account = accounts.FirstOrDefault();
        try
        {
            var result = await _pca.AcquireTokenSilent(Scopes, account).ExecuteAsync(ct);
            SignedInUser = result.Account.Username;
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            if (!interactiveAllowed)
                throw new AuthRequiredException("Sign-in required. Open SendToOneNote and sign in.");
            try
            {
                var result = await _pca.AcquireTokenInteractive(Scopes).ExecuteAsync(ct);
                SignedInUser = result.Account.Username;
                return result.AccessToken;
            }
            catch (MsalException ex)
            {
                throw new AuthRequiredException($"Sign-in was cancelled or failed: {ex.Message}");
            }
        }
    }

    private async Task EnsureCacheAsync()
    {
        if (_cacheAttached) return;
        await _init.WaitAsync();
        try
        {
            if (_cacheAttached) return;
            var props = new StorageCreationPropertiesBuilder("msal_cache.bin", _cacheDir).Build();
            var helper = await MsalCacheHelper.CreateAsync(props);
            helper.RegisterCache(_pca.UserTokenCache);
            _cacheAttached = true;
        }
        finally { _init.Release(); }
    }
}
