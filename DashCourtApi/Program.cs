using DashCourtApi.Services;
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Configure CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigin",
        builder => builder.WithOrigins("http://localhost:5500", "http://localhost:8000") // Replace with your frontend URL
                            .AllowAnyHeader()
                            .AllowAnyMethod());
});

// Register ExcelDataService as a singleton
string excelFilesPath = Path.Combine(Directory.GetCurrentDirectory(), "XLSX"); // Path to the XLSX files
string jsonOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "XLSX"); // Path to save JSON files in the XLSX directory
Console.WriteLine($"JSON output path: {jsonOutputPath}");
bool useMockData = builder.Configuration.GetValue<bool>("AppSettings:UseMockData");
var excelDataService = new ExcelDataService(excelFilesPath, useMockData, jsonOutputPath);
builder.Services.AddSingleton(excelDataService);

// Generate JSON files if not using mock data
if (!useMockData)
{
    excelDataService.GenerateJsonFiles();
}

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

// Use CORS
app.UseCors("AllowSpecificOrigin");

app.UseAuthorization();

app.MapControllers();

app.Run();
