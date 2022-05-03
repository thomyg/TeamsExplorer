
var builder = Microsoft.AspNetCore.Builder.WebApplication.CreateBuilder(args);

builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddTeamsFx(builder.Configuration.GetSection("TeamsFx"));
builder.Services.AddScoped<MicrosoftTeams>();

//Service for Radzen Notifications
builder.Services.AddScoped<NotificationService>();

// Add and configure PnPCore and PnPCoreAuth services
// Check out the appsettings.json for the configuration details
builder.Services.AddPnPCore();
builder.Services.Configure<PnPCoreOptions>(builder.Configuration.GetSection("PnPCore"));
builder.Services.AddPnPCoreAuthentication();
builder.Services.Configure<PnPCoreAuthenticationOptions>(builder.Configuration.GetSection("PnPCore"));

// Add our own service for calling Graph and PnPCore
builder.Services.AddScoped<DataService>();

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.UseEndpoints(endpoints =>
{
    endpoints.MapBlazorHub();
    endpoints.MapFallbackToPage("/_Host");
    endpoints.MapControllers();
});

app.Run();

