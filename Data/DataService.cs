using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using System.Linq.Expressions;

namespace TeamsExplorer.Data;

public class DataService
{    
    TeamsUserCredential teamsUserCredential;
    MicrosoftTeams microsoftTeams;    
    IPnPContextFactory pnPContextFactory;    

    private readonly string _scope = "User.Read";

    GraphServiceClient GraphClient;

    public DataService(TeamsUserCredential teamsUserCredential, 
         MicrosoftTeams microsoftTeams, IPnPContextFactory PnPContextFactory)
    {        
        this.teamsUserCredential = teamsUserCredential;        
        this.microsoftTeams = microsoftTeams;
        this.pnPContextFactory = PnPContextFactory;        
    }
    private async Task<GraphServiceClient> GetGraphServiceClient()
    {
        await teamsUserCredential.GetTokenAsync(new TokenRequestContext(new string[] { _scope }), new System.Threading.CancellationToken());
        var msGraphAuthProvider = new MsGraphAuthProvider(teamsUserCredential, _scope);
        var client = new GraphServiceClient(msGraphAuthProvider);
        return client;
    }
    public async Task<List<Team>> GetTeams()
    {
        GraphClient = await GetGraphServiceClient();

        var teams = await GraphClient
                    .Me
                    .JoinedTeams
                    .Request()
                    .GetAsync();

        //For demo purpose we only care about the first returned page
        //If you have a multitude of joined teams add paging
        List<Team> result = (List<Team>)teams.CurrentPage;
        result.Sort((a, b) => a.DisplayName.CompareTo(b.DisplayName));

        return result;
    }    
    public async Task<Dictionary<string, string>> GetTeamDetails(string teamId)
    {
        Dictionary<string, string> result = new();

        var team = await GraphClient
                .Teams[teamId]
                .Request()
                .GetAsync();

        foreach (PropertyInfo p in team.GetType().GetProperties())
        {
            var value = p.GetValue(team, null);

            if (value != null)
                result.Add(p.Name, JsonSerializer.Serialize(value));

            if (value == null)
                result.Add(p.Name, "null");
        }

        var teamSite = await GraphClient
                    .Groups[teamId]
                    .Sites["root"]
                    .Request()
                    .GetAsync();

        result.Add("SPSiteUrl", teamSite.WebUrl);

        result = result.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        return result;
    }

    public async Task<Dictionary<string, string>> GetInstalledApps(string teamId)
    {
        Dictionary<string, string> result = new();

        var apps = await GraphClient
                .Teams[teamId].InstalledApps
                .Request()
                .Expand("teamsAppDefinition")
                .GetAsync();

        foreach (var app in apps.CurrentPage)
        {
            string name = app.TeamsAppDefinition.DisplayName + " - " + app.TeamsAppDefinition.TeamsAppId;
            result.Add(name, JsonSerializer.Serialize(app));
        }

        result = result.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
        return result;
    }

    public async Task<Dictionary<string, string>> GetContextProperties()
    {
        Dictionary<string, string> result = new();

        TeamsContext ctx = await microsoftTeams.GetTeamsContextAsync();

        foreach (PropertyInfo p in ctx.GetType().GetProperties())
        {
            var value = p.GetValue(ctx, null);

            if (value != null)
                result.Add(p.Name, JsonSerializer.Serialize(value));

            if (value == null)
                result.Add(p.Name, "null");
        }

        return result;
    }

    public async Task<Dictionary<string, string>> GetSiteProperties(string teamId)
    {
        Dictionary<string, string> result = new();

        Site root = await GraphClient.Groups[teamId]
                        .Sites["root"]
                        .Request()
                        .GetAsync();

        var options = new PnPContextOptions()
        {
            AdditionalSitePropertiesOnCreate =
                new Expression<Func<ISite, object>>[] {
                        s => s.Url,
                        s => s.HubSiteId,
                        s => s.Features,
                        s => s.All
            },
            AdditionalWebPropertiesOnCreate =
                new Expression<Func<IWeb, object>>[]
                {
                        w => w.ServerRelativeUrl,
                        w => w.Fields,
                        w => w.Features,
                        w => w.Lists.QueryProperties(r => r.Title,
                            r => r.RootFolder.QueryProperties(p => p.ServerRelativeUrl)),
                        w => w.AllProperties
                }
        };

        using (var context = await pnPContextFactory.CreateAsync(new Uri(root.WebUrl), options))
        {
            var webId = context.Web.Id;
            await context.Web.LoadAsync();
            await context.Site.LoadAsync();

            foreach (PropertyInfo p in context.Site.GetType().GetProperties())
            {
                try
                {
                    var value = p.GetValue(context.Site, null);

                    if (value != null)
                        result.Add(p.Name, JsonSerializer.Serialize(value));

                    if (value == null)
                        result.Add(p.Name, "null");
                }
                catch (Exception ex)
                {
                    //As a lot of properties are not set, we catch each exception here
                    //and set the value in the result to null
                    result.Add(p.Name, "null");
                }
            }

            int x = 10;
        };

        return result;
    }

    public async Task<Dictionary<string, string>> GetWebProperties(string teamId)
    {
        Dictionary<string, string> result = new();

        Site root = await GraphClient.Groups[teamId]
                        .Sites["root"]
                        .Request()
                        .GetAsync();

        var options = new PnPContextOptions()
        {
            AdditionalSitePropertiesOnCreate =
                new Expression<Func<ISite, object>>[] {
                        s => s.Url,
                        s => s.HubSiteId,
                        s => s.Features,
                        s => s.All
            },
            AdditionalWebPropertiesOnCreate =
                new Expression<Func<IWeb, object>>[]
                {
                        w => w.ServerRelativeUrl,
                        w => w.Fields,
                        w => w.Features,
                        w => w.Lists.QueryProperties(r => r.Title,
                            r => r.RootFolder.QueryProperties(p => p.ServerRelativeUrl)),
                        w => w.AllProperties
                }
        };

        using (var context = await pnPContextFactory.CreateAsync(new Uri(root.WebUrl), options))
        {
            var webId = context.Web.Id;
            await context.Web.LoadAsync();
            await context.Site.LoadAsync();

            var values = context.Web.AllProperties.Values;
            foreach (var property in values)
            {
                result.Add(property.Key, (string)property.Value);
            }

            foreach (PropertyInfo p in context.Web.GetType().GetProperties())
            {
                try
                {
                    var value = p.GetValue(context.Web, null);

                    if (value != null)
                        result.Add(p.Name, JsonSerializer.Serialize(value));

                    if (value == null)
                        result.Add(p.Name, "null");
                }
                catch (Exception ex)
                {
                    //As a lot of properties are not set, we catch each exception here
                    //and set the value in the result to null
                    result.Add(p.Name, "null");
                }
            }
        };

        return result;
    }

    public async Task<List<AadUserConversationMember>> GetMembers(string teamId)
    {
        List<AadUserConversationMember> result = new();

        var members = await GraphClient
                .Teams[teamId].Members
                .Request()
                .GetAsync();

        foreach (AadUserConversationMember member in members.CurrentPage)
        {
            result.Add(member);
        }

        return result;
    }

    public async Task<List<Channel>> GetChannels(string teamdId)
    {
        List<Channel> result = new();

        var channels = await GraphClient
            .Teams[teamdId]
            .Channels
            .Request()
            .GetAsync();

        result = channels.CurrentPage.ToList<Channel>();
        result.Sort((a, b) => a.DisplayName.CompareTo(b.DisplayName));

        return result;
    }

    public async Task<Dictionary<string, List<TeamsTab>>> GetTabs(string teamId, List<Channel> channels)
    {
        Dictionary<string, List<TeamsTab>> result = new();

        foreach (Channel c in channels)
        {
            var tabs = await GraphClient
                .Teams[teamId]
                .Channels[c.Id]
                .Tabs
                .Request()
                .GetAsync();

            result.Add(c.Id, tabs.CurrentPage.ToList());
        }

        return result;
    }

}
