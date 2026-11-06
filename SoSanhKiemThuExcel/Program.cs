using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;
using SoSanhKiemThuExcel.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

builder.Services.AddScoped<IExcelCompareService, ExcelCompareService>();

// FIX: Set EPPlus license context (NonCommercial cho miễn phí)
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Cho phép upload file lớn hơn
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 100 * 1024 * 1024; // 100MB
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
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
