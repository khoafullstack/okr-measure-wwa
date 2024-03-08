using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using OKR.DORA.Models;
using OKR.DORA.Services;
using System.Text.Json;

namespace OKR.DORA.Controllers;
public class EmbedInfoController : Controller
{
    private readonly PbiEmbedService pbiEmbedService;
    private readonly IOptions<AzureAd> azureAd;
    private readonly IOptions<PowerBI> powerBI;

    public EmbedInfoController(PbiEmbedService pbiEmbedService, IOptions<AzureAd> azureAd, IOptions<PowerBI> powerBI)
    {
        this.pbiEmbedService = pbiEmbedService;
        this.azureAd = azureAd;
        this.powerBI = powerBI;
    }

    /// <summary>
    /// Returns Embed token, Embed URL, and Embed token expiry to the client
    /// </summary>
    /// <returns>JSON containing parameters for embedding</returns>
    [HttpGet]
    public string GetEmbedInfo()
    {
        try
        {
            // Validate whether all the required configurations are provided in appsettings.json
            string configValidationResult = ConfigValidatorService.ValidateConfig(azureAd, powerBI);
            if (configValidationResult != null)
            {
                HttpContext.Response.StatusCode = 400;
                return configValidationResult;
            }

            EmbedParams embedParams = pbiEmbedService.GetEmbedParams(new Guid(powerBI.Value.WorkspaceId), new Guid(powerBI.Value.ReportId));
            return JsonSerializer.Serialize<EmbedParams>(embedParams);
        }
        catch (Exception ex)
        {
            HttpContext.Response.StatusCode = 500;
            return ex.Message + "\n\n" + ex.StackTrace;
        }
    }
}
