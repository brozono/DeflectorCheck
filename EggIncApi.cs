namespace EggIncApi;

using Ei;
using Google.Protobuf;

#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
public class EggIncApi
{
    private const int CLIENTVERSION = 40;

    public static async Task<ContractCoopStatusResponse> GetCoopStatus(string contractId, string coopId, string userId)
    {
        ContractCoopStatusRequest coopStatusRequest = new ()
        {
            ContractIdentifier = contractId,
            CoopIdentifier = coopId,
            UserId = userId,
        };

        return await MakeEggIncApiRequest("coop_status", coopStatusRequest, ContractCoopStatusResponse.Parser.ParseFrom);
    }

    public static async Task<EggIncFirstContactResponse> GetFirstContact(string userId)
    {
        EggIncFirstContactRequest firstContactRequest = new ()
        {
            EiUserId = userId,
            ClientVersion = CLIENTVERSION,
        };

        return await MakeEggIncApiRequest("bot_first_contact", firstContactRequest, EggIncFirstContactResponse.Parser.ParseFrom, false);
    }

    public static async Task<PeriodicalsResponse> GetPeriodicals(string userId)
    {
        GetPeriodicalsRequest getPeriodicalsRequest = new ()
        {
            UserId = userId,
            CurrentClientVersion = CLIENTVERSION,
        };

        return await MakeEggIncApiRequest("get_periodicals", getPeriodicalsRequest, PeriodicalsResponse.Parser.ParseFrom);
    }

    private static async Task<T> MakeEggIncApiRequest<T>(string endpoint, IMessage data, Func<ByteString, T> parseMethod, bool authenticated = true)
    {
        byte[] bytes;
        using (var stream = new MemoryStream())
        {
            data.WriteTo(stream);
            bytes = stream.ToArray();
        }

        string response = await PostRequest($"https://www.auxbrain.com/ei/{endpoint}", new FormUrlEncodedContent(new Dictionary<string, string>
        {
            { "data", Convert.ToBase64String(bytes) },
        }));

        if (authenticated)
        {
            AuthenticatedMessage authenticatedMessage = AuthenticatedMessage.Parser.ParseFrom(Convert.FromBase64String(response));
            return parseMethod(authenticatedMessage.Message);
        }
        else
        {
            return parseMethod(ByteString.CopyFrom(Convert.FromBase64String(response)));
        }
    }

    private static async Task<string> PostRequest(string url, FormUrlEncodedContent json)
    {
        using var client = new HttpClient();
        var response = await client.PostAsync(url, json);
        return await response.Content.ReadAsStringAsync();
    }
}
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member