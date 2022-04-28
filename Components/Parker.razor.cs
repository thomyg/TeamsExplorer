using Microsoft.Extensions.Configuration;
using System;

namespace TeamsExplorer.Components
{
    public partial class Parker
    {
        
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
}
