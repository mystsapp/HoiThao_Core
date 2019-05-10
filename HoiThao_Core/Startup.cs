using Data;
using Data.Repository;
using DinkToPdf;
using DinkToPdf.Contracts;
using HoiThao_Core.Helpers;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using Wkhtmltopdf.NetCore;

namespace HoiThao_Core
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.Configure<CookiePolicyOptions>(options =>
            {
                // This lambda determines whether user consent for non-essential cookies is needed for a given request.
                options.CheckConsentNeeded = context => true;
                options.MinimumSameSitePolicy = SameSiteMode.None;
            });

            services.AddDbContext<hoinghiContext>(options =>
        options.UseSqlServer(Configuration.GetConnectionString("DefaultConnection")));

            //services.AddScoped<HoiNghiDbContext>(_ => new HoiNghiDbContext(Configuration.GetConnectionString("DbConnection")));

            services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
            var contextDll = new CustomAssemblyLoadContext();
            contextDll.LoadUnmanagedLibrary(Path.Combine(Directory.GetCurrentDirectory(), "libwkhtmltox.dll"));

            services.AddTransient<IAccountRepository, AccountRepository>();
            services.AddTransient<IAseanRepository, AseanRepository>();

            //services.AddSingleton<IFileProvider>(
            //new PhysicalFileProvider(
            //    Path.Combine(Directory.GetCurrentDirectory(), "wwwroot")));
            //services.AddScoped(provider =>
            //{
            //    //var connectionString = Configuration.GetConnectionString("DefaultConnection");
            //    return new HoiNghiDbContext("metadata=res://*/Context.HoiNghiDbContext.csdl|res://*/Context.HoiNghiDbContext.ssdl|res://*/Context.HoiNghiDbContext.msl;provider=System.Data.SqlClient;provider connection string=&quot;data source=192.168.4.198;initial catalog=hoinghi;user id=sa;password=TigerSts@2017;MultipleActiveResultSets=True;App=EntityFramework&quot;");
            //});
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
            }

            app.UseStaticFiles();
            app.UseCookiePolicy();

            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");
            });

            //Rotativa.AspNetCore.RotativaConfiguration.Setup(env);
            Rotativa.AspNetCore.RotativaConfiguration.Setup(env);
        }
        
    }
}