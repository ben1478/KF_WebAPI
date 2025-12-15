using KF_WebAPI.Controllers;
using KF_WebAPI.Service;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

// Excel 授權設定
ExcelPackage.LicenseContext = LicenseContext.Commercial;

// 加入服務
builder.Services.AddControllers();
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddHttpClient<YuRichAPIController>();
builder.Services.AddScoped<IWebRobotService, WebRobotService>();



// CORS 設定
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowFrontendOrigins", policy =>
    {
        policy.WithOrigins(
            "https://kuofongserver.ngrok.pro",
            "https://kfserver-jwt.ngrok.pro",
            "http://192.168.1.240:8080",
            "http://192.168.1.240:8081",
            "http://localhost:5000"
        )
        .AllowAnyMethod()
        .AllowAnyHeader()
        .WithExposedHeaders("Content-Disposition")
        .AllowCredentials();
    });
});

var app = builder.Build();

// Swagger (僅在開發環境)
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

// Session
app.UseSession();

// CORS 必須在 Routing 後、Auth 前
app.UseCors("AllowFrontendOrigins");

app.UseAuthentication();
app.UseAuthorization();



// 套用 CORS 到所有 Controller
app.MapControllers().RequireCors("AllowFrontendOrigins");

app.Run();