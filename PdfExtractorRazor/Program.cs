using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PdfExtractorRazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorPages();

builder.Services.AddScoped<IPdfExtractionService, CompletePdfExtractionService>();

//builder.WebHost.UseWebRoot("wwwroot");              // sets WebRootPath
var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
	app.UseExceptionHandler("/Error");
	app.UseHsts();
}

//// Configure file upload limits
//builder.Services.Configure<IISServerOptions>(options =>
//{
//    options.MaxRequestBodySize = 100 * 1024 * 1024; // 100 MB
//});
app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.MapRazorPages(); 


app.Run();

  
