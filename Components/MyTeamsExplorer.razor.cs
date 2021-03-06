
namespace TeamsExplorer.Components;

public partial class MyTeamsExplorer
{
    string errorMessage = "";
    bool isInTeams;        
    string SelectedTeamId;
    bool showSPOData = false;        
    bool noTeamSelected = true;

    [Inject]
    NotificationService NotificationService { get; set; }
    [Inject]
    MicrosoftTeams MicrosoftTeams { get; set; }
    [Inject]
    IWebHostEnvironment HostEnvironment { get; set; }
    [Inject]
    IConfiguration Configuration { get; set; }
    [Inject]
    DataService DataService { get; set; }

    List<Team> JoinedTeams = new List<Team>();
    Dictionary<string, string> SelectedTeamProps;
    Dictionary<string, string> InstalledApps;
    List<AadUserConversationMember> Members;
    List<Channel> Channels;
    Dictionary<string, List<TeamsTab>> Tabs;
    Dictionary<string, string> ContextProperties;
    Dictionary<string, string> SiteProperties;
    Dictionary<string, string> WebProperties;

    protected override void OnInitialized()
    {
        base.OnInitialized();
        // Get configuration value to modify HTML in order to include SharePoint data or not
        showSPOData = Convert.ToBoolean(Configuration["TeamsExplorer:ShowSPOData"]);
    }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {

        await base.OnAfterRenderAsync(firstRender);

        try
        {
            if (firstRender)
            {
                isInTeams = await MicrosoftTeams.IsInTeams();

                // Get joined teams of the user and load them into the dropdown
                // From here the dropdown drives the app behaviour with its change events
                JoinedTeams = await DataService.GetTeams();
                ShowNotification(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Teams Summary", Detail = "Successfully got data from " + JoinedTeams.Count + " teams.", Duration = 4000 });
                StateHasChanged();
            }
        }
        catch (Exception ex)
        {
            errorMessage = ex.ToString();
            ShowNotification(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error Summary", Detail = errorMessage, Duration = 8000 });
            StateHasChanged();
        }
    }

    private async void OnTeamsSelectChange(Object team)
    {
        try
        {
            noTeamSelected = false;            
            ClearPropertyLists();

            Team selectedTeam = (Team)team;
            SelectedTeamId = selectedTeam.Id;

            // Getting all the Graph data
            SelectedTeamProps = await DataService.GetTeamDetails(SelectedTeamId);
            InstalledApps = await DataService.GetInstalledApps(SelectedTeamId);
            Members = await DataService.GetMembers(SelectedTeamId);
            Channels = await DataService.GetChannels(SelectedTeamId);
            Tabs = await DataService.GetTabs(SelectedTeamId, Channels);         
            ContextProperties = await DataService.GetContextProperties();

            StateHasChanged();

            if (showSPOData)
            {
                // If configuration is showSPOData=true also query SPO properties
                SiteProperties = await DataService.GetSiteProperties(SelectedTeamId);                
                WebProperties = await DataService.GetWebProperties(SelectedTeamId);
                StateHasChanged();
            }

            ShowNotification(new NotificationMessage { Severity = NotificationSeverity.Success, Summary = "Data Request", Detail = "Successfully got data for the selected team.", Duration = 2000 });
        }
        catch (Exception ex)
        {
            errorMessage = ex.ToString();
            ShowNotification(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error Summary", Detail = errorMessage, Duration = 8000 });
            StateHasChanged();
        }
    }
            
    private void ClearPropertyLists()
    {
        if(SelectedTeamProps!=null)
            SelectedTeamProps.Clear();

        if(InstalledApps!=null)
            InstalledApps.Clear();
        
        if(ContextProperties!=null)
            ContextProperties.Clear();
        
        if(SiteProperties!=null)
            SiteProperties.Clear();
        
        if(WebProperties!=null)
            WebProperties.Clear();

        if(Members!=null)
            Members.Clear();
        
        if(Channels!=null)
            Channels.Clear();
        
        if(Tabs!=null)
            Tabs.Clear();            

        StateHasChanged();
    }

    void ShowNotification(NotificationMessage message)
    {
        NotificationService.Notify(message);
    }
}
