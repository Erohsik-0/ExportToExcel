using CsvExporterLibrary.DIContainer;
using ExportExcel.DependencyInjectionContainer;
using PresentationExporter.DependencyInjectionContainer;
using TimeZoneConvertorLibrary.Extensions;

var builder = WebApplication.CreateBuilder(args);


// Add services to the container.
builder.Services.AddControllersWithViews();


//Registering the services for CsvExportPkg library
builder.Services.AddCsvExporterServices();

//Registering the services for TimeZoneConvertor Library
builder.Services.AddTimeZoneConversionServices();

//Registering the services for ExportExcel Library
builder.Services.AddExcelExporterServices();

//Registering the services for PresentationExporter Library
builder.Services.AddPresentationExportServices();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();

app.UseAuthorization();

app.MapControllers();

app.MapStaticAssets();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}")
    .WithStaticAssets();


app.Run();
