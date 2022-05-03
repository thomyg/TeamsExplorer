
namespace TeamsExplorer.Components;

public partial class ChannelTable
{
    [Parameter]
    public List<Channel> Channels { get; set; }

    [Parameter]
    public Dictionary<string, List<TeamsTab>> Tabs { get; set; }
}
