namespace SendToOneNote.Core.Auth;

public interface ITokenProvider
{
    Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default);
    string? SignedInUser { get; }
}

public sealed class AuthRequiredException(string message) : Exception(message);
