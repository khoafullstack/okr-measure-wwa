using Microsoft.PowerBI.Api.Models;
using Microsoft.PowerBI.Api;
using Microsoft.Rest;
using OKR.DORA.Models;
using System.Runtime.InteropServices;

namespace OKR.DORA.Services;

public class PbiEmbedService
{
    private readonly AadService aadService;
    private readonly string powerBiApiUrl = "https://api.powerbi.com";
    private readonly string _token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSIsImtpZCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvYjBiNTliOTgtYTcxZC00ZDIzLWFhM2QtMjBiMjRiZThkMjJhLyIsImlhdCI6MTcwOTg2ODU3NiwibmJmIjoxNzA5ODY4NTc2LCJleHAiOjE3MDk4NzM0MDUsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVFFBeS84V0FBQUFXZVV3N05zcktDS2pGb3M5aUtIWGZZSXlvaDByMXNIeHpTQVBnamhTRUNEcWF1M1UrYlVoVnlocmR1cDJ1Q0o5IiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjIzZDhmNmJkLTFlYjAtNGNjMi1hMDhjLTdiZjUyNWM2N2JjZCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiTmd1eWVuIiwiZ2l2ZW5fbmFtZSI6IlBodSIsImlwYWRkciI6IjExNS43OS41LjE5NyIsIm5hbWUiOiJQaHUgTmd1eWVuIiwib2lkIjoiNzQ5YzQ5N2YtZGZiNy00NDg5LWE1NmItMmNjMWNkMzI3NzM2IiwicHVpZCI6IjEwMDMyMDAzNUU3OUFCRUMiLCJyaCI6IjAuQVZVQW1KdTFzQjJuSTAycVBTQ3lTLWpTS2drQUFBQUFBQUFBd0FBQUFBQUFBQUNfQVBJLiIsInNjcCI6IkFwcC5SZWFkLkFsbCBDYXBhY2l0eS5SZWFkLkFsbCBDYXBhY2l0eS5SZWFkV3JpdGUuQWxsIENvbnRlbnQuQ3JlYXRlIERhc2hib2FyZC5SZWFkLkFsbCBEYXNoYm9hcmQuUmVhZFdyaXRlLkFsbCBEYXRhZmxvdy5SZWFkLkFsbCBEYXRhZmxvdy5SZWFkV3JpdGUuQWxsIERhdGFzZXQuUmVhZC5BbGwgRGF0YXNldC5SZWFkV3JpdGUuQWxsIEdhdGV3YXkuUmVhZC5BbGwgR2F0ZXdheS5SZWFkV3JpdGUuQWxsIFBpcGVsaW5lLkRlcGxveSBQaXBlbGluZS5SZWFkLkFsbCBQaXBlbGluZS5SZWFkV3JpdGUuQWxsIFJlcG9ydC5SZWFkLkFsbCBSZXBvcnQuUmVhZFdyaXRlLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkV3JpdGUuQWxsIFRlbmFudC5SZWFkLkFsbCBUZW5hbnQuUmVhZFdyaXRlLkFsbCBVc2VyU3RhdGUuUmVhZFdyaXRlLkFsbCBXb3Jrc3BhY2UuUmVhZC5BbGwgV29ya3NwYWNlLlJlYWRXcml0ZS5BbGwiLCJzdWIiOiJ4dlc2TjkyRHNvVS1OOHpQWEl5aE12ZU9TczZvcDJWRXdzWGp5MExRNXZJIiwidGlkIjoiYjBiNTliOTgtYTcxZC00ZDIzLWFhM2QtMjBiMjRiZThkMjJhIiwidW5pcXVlX25hbWUiOiJwaHUubmd1eWVuQGV6dGVrLnZuIiwidXBuIjoicGh1Lm5ndXllbkBlenRlay52biIsInV0aSI6IlprVlFxWUNnVDBxeHI4dkRwdFNuQVEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdfQ.AnjFjNcFbEDUyKFqjnRJ4nP9DqmVqZp_X0dOXKvuGinmD_S2U38zDJ_AY1vYBX_afdP444MlWkqQZovCfF-EFeCd2xY0gPuMBYgcJb1KefRmGjaiZltiF4XE6pvaqoM4Nm3Og5NODQ6UwNV5FiaKKy2PtdX8F_EU25wieNZcAKyWIfr2AtXnhXyIENphUutgNrYA8iltKR1ihP7BJCAGAkqu1_UU1iD8cReiTrjv2dhv6g8UXrnQz_wMZhHl2Jq3CIaJKCHhXxqVgHRNJsfEm8APNHhZk2lZ-PHzN4UaCtktuANHE4An2HfBw2Lh6xIHwxmH-ODY-jY4_QU9S3xuVA";

    public PbiEmbedService(AadService aadService)
    {
        this.aadService = aadService;
    }

    /// <summary>
    /// Get Power BI client
    /// </summary>
    /// <returns>Power BI client object</returns>
    public PowerBIClient GetPowerBIClient()
    {
        //var tokenCredentials = new TokenCredentials(aadService.GetAccessToken(), "Bearer");
        var tokenCredentials = new TokenCredentials(_token, "Bearer");
        return new PowerBIClient(new Uri(powerBiApiUrl), tokenCredentials);
    }

    /// <summary>
    /// Get embed params for a report
    /// </summary>
    /// <returns>Wrapper object containing Embed token, Embed URL, Report Id, and Report name for single report</returns>
    public EmbedParams GetEmbedParams(Guid workspaceId, Guid reportId, [Optional] Guid additionalDatasetId)
    {
        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Get report info
        var pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);

        //  Check if dataset is present for the corresponding report
        //  If isRDLReport is true then it is a RDL Report 
        var isRDLReport = String.IsNullOrEmpty(pbiReport.DatasetId);

        EmbedToken embedToken;

        // Generate embed token for RDL report if dataset is not present
        if (isRDLReport)
        {
            // Get Embed token for RDL Report
            embedToken = GetEmbedTokenForRDLReport(workspaceId, reportId);
        }
        else
        {
            // Create list of datasets
            var datasetIds = new List<Guid>();

            // Add dataset associated to the report
            datasetIds.Add(Guid.Parse(pbiReport.DatasetId));

            // Append additional dataset to the list to achieve dynamic binding later
            if (additionalDatasetId != Guid.Empty)
            {
                datasetIds.Add(additionalDatasetId);
            }

            // Get Embed token multiple resources
            embedToken = GetEmbedToken(reportId, datasetIds, workspaceId);
        }

        // Add report data for embedding
        var embedReports = new List<EmbedReport>() {
                new EmbedReport
                {
                    ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl
                }
            };

        // Capture embed params
        var embedParams = new EmbedParams
        {
            EmbedReport = embedReports,
            Type = "Report",
            EmbedToken = embedToken
        };

        return embedParams;
    }

    /// <summary>
    /// Get embed params for multiple reports for a single workspace
    /// </summary>
    /// <returns>Wrapper object containing Embed token, Embed URL, Report Id, and Report name for multiple reports</returns>
    /// <remarks>This function is not supported for RDL Report</remakrs>
    public EmbedParams GetEmbedParams(Guid workspaceId, IList<Guid> reportIds, [Optional] IList<Guid> additionalDatasetIds)
    {
        // Note: This method is an example and is not consumed in this sample app

        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Create mapping for reports and Embed URLs
        var embedReports = new List<EmbedReport>();

        // Create list of datasets
        var datasetIds = new List<Guid>();

        // Get datasets and Embed URLs for all the reports
        foreach (var reportId in reportIds)
        {
            // Get report info
            var pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);

            datasetIds.Add(Guid.Parse(pbiReport.DatasetId));

            // Add report data for embedding
            embedReports.Add(new EmbedReport { ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl });
        }

        // Append to existing list of datasets to achieve dynamic binding later
        if (additionalDatasetIds != null)
        {
            datasetIds.AddRange(additionalDatasetIds);
        }

        // Get Embed token multiple resources
        var embedToken = GetEmbedToken(reportIds, datasetIds, workspaceId);

        // Capture embed params
        var embedParams = new EmbedParams
        {
            EmbedReport = embedReports,
            Type = "Report",
            EmbedToken = embedToken
        };

        return embedParams;
    }

    /// <summary>
    /// Get Embed token for single report, multiple datasets, and an optional target workspace
    /// </summary>
    /// <returns>Embed token</returns>
    /// <remarks>This function is not supported for RDL Report</remakrs>
    public EmbedToken GetEmbedToken(Guid reportId, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
    {
        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Create a request for getting Embed token 
        // This method works only with new Power BI V2 workspace experience
        var tokenRequest = new GenerateTokenRequestV2(

            reports: new List<GenerateTokenRequestV2Report>() { new GenerateTokenRequestV2Report(reportId) },

            datasets: datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList(),

            targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
        );

        // Generate Embed token
        var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

        return embedToken;
    }

    /// <summary>
    /// Get Embed token for multiple reports, datasets, and an optional target workspace
    /// </summary>
    /// <returns>Embed token</returns>
    /// <remarks>This function is not supported for RDL Report</remakrs>
    public EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
    {
        // Note: This method is an example and is not consumed in this sample app

        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Convert report Ids to required types
        var reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

        // Convert dataset Ids to required types
        var datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

        // Create a request for getting Embed token 
        // This method works only with new Power BI V2 workspace experience
        var tokenRequest = new GenerateTokenRequestV2(

            datasets: datasets,

            reports: reports,

            targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
        );

        // Generate Embed token
        var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

        return embedToken;
    }

    /// <summary>
    /// Get Embed token for multiple reports, datasets, and optional target workspaces
    /// </summary>
    /// <returns>Embed token</returns>
    /// <remarks>This function is not supported for RDL Report</remakrs>
    public EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] IList<Guid> targetWorkspaceIds)
    {
        // Note: This method is an example and is not consumed in this sample app

        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Convert report Ids to required types
        var reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

        // Convert dataset Ids to required types
        var datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

        // Convert target workspace Ids to required types
        IList<GenerateTokenRequestV2TargetWorkspace> targetWorkspaces = null;
        if (targetWorkspaceIds != null)
        {
            targetWorkspaces = targetWorkspaceIds.Select(targetWorkspaceId => new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId)).ToList();
        }

        // Create a request for getting Embed token 
        // This method works only with new Power BI V2 workspace experience
        var tokenRequest = new GenerateTokenRequestV2(

            datasets: datasets,

            reports: reports,

            targetWorkspaces: targetWorkspaceIds != null ? targetWorkspaces : null
        );

        // Generate Embed token
        var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

        return embedToken;
    }

    /// <summary>
    /// Get Embed token for RDL Report
    /// </summary>
    /// <returns>Embed token</returns>
    public EmbedToken GetEmbedTokenForRDLReport(Guid targetWorkspaceId, Guid reportId, string accessLevel = "view")
    {
        PowerBIClient pbiClient = this.GetPowerBIClient();

        // Generate token request for RDL Report
        var generateTokenRequestParameters = new GenerateTokenRequest(
            accessLevel: accessLevel
        );

        // Generate Embed token
        var embedToken = pbiClient.Reports.GenerateTokenInGroup(targetWorkspaceId, reportId, generateTokenRequestParameters);

        return embedToken;
    }
}
