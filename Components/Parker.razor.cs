
// A very simple class to showcase how you could
// structure your Blazor solution by using razor components.

namespace TeamsExplorer.Components;

public partial class Parker
{
    [Inject]
    IConfiguration Configuration { get; set; }

    bool fancyMode = false;

    public Parker()
    {

    }
    protected override void OnInitialized()
    {
        base.OnInitialized();
        fancyMode = Convert.ToBoolean(Configuration["TeamsExplorer:FancyMode"]);
    }
}
