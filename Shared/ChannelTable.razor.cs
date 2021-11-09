using Microsoft.AspNetCore.Components;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeamsExplorer.Shared
{
    public partial class ChannelTable
    {
        [Parameter]
        public List<Channel> Channels { get; set; }

        [Parameter]
        public Dictionary<string, List<TeamsTab>> Tabs { get; set; }
    }
}
