using KF_WebAPI.Controllers;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

// �]�w EPPlus ���v�W�U��
ExcelPackage.LicenseContext = LicenseContext.Commercial;
// Add services to the container.

builder.Services.AddControllers();
/*builder.Services.AddControllers().AddJsonOptions(x =>
{
    x.JsonSerializerOptions.NumberHandling = JsonNumberHandling.AllowReadingFromString;
    x.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
});*/

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle

// Add Session services
builder.Services.AddDistributedMemoryCache(); 
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromHours(8); // �]�m Session �L���ɶ�
    options.Cookie.HttpOnly = true; 
    options.Cookie.IsEssential = true; 
});
builder.Services.AddControllers();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddHttpClient<YuRichAPIController>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
else 
{

        app.UseExceptionHandler("/Error");
        // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
        app.UseHsts();

}


app.UseCors(builder =>
{
    builder.WithOrigins(
        "http://erp",
        "http://192.168.1.240",
        "http://192.168.1.27/KF_Web/",
        "http://192.168.1.27/KF_WebAPI/",
        "https://www.kuofongweb.com.tw/"
        )
           .AllowAnyMethod()
           .AllowAnyHeader()
           .AllowCredentials(); // ���\�o�e���ҡ]�]�A Cookies�^
});

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

// �ҥ� Session �䴩
app.UseSession();

app.UseAuthorization();

app.MapControllers();

app.Run();
