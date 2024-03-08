using Microsoft.Extensions.Configuration;
using OKR.DORA.Models;
using OKR.DORA.Services;

namespace OKR.DORA;

public class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        builder.Services.AddScoped(typeof(AadService))
                    .AddScoped(typeof(PbiEmbedService));

        // Add services to the container.
        builder.Services.AddControllersWithViews();

        builder.Services.Configure<AzureAd>(builder.Configuration.GetSection("AzureAd"))
                    .Configure<PowerBI>(builder.Configuration.GetSection("PowerBI"));

        var app = builder.Build();

        // Configure the HTTP request pipeline.
        if (!app.Environment.IsDevelopment())
        {
            app.UseExceptionHandler("/Home/Error");
            // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
            app.UseHsts();
        }

        app.UseHttpsRedirection();
        app.UseStaticFiles();

        app.UseRouting();

        app.UseAuthorization();

        app.MapControllerRoute(
            name: "default",
            pattern: "{controller=Home}/{action=Index}/{id?}");

        app.Run();
    }
}
