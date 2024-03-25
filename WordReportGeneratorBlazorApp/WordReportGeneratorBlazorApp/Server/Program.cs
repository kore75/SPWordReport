using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.Identity.Web;

namespace WordReportGeneratorBlazorApp
{
    //https://www.reddit.com/user/azzurrabrancati/comments/f1r3op/consume_azure_ad_secured_web_api_net_core_31_from/

    /// <summary>
    /// Consume Azure AD secured Web API .NET Core 3.1 from your SPFx code
    /// </summary>
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddMicrosoftIdentityWebApi(builder.Configuration.GetSection("AzureAd"));

            builder.Services.AddControllersWithViews();
            builder.Services.AddRazorPages();
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("CorsPolicy", builder => builder.WithOrigins("https://devkoeli.sharepoint.com")
                .AllowAnyMethod()
                .AllowAnyHeader());
            });

            var app = builder.Build();
            app.UseCors("CorsPolicy");

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseWebAssemblyDebugging();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            app.UseBlazorFrameworkFiles();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();


            app.MapRazorPages();
            app.MapControllers();
            app.MapFallbackToFile("index.html");

            app.Run();
        }
    }
}
