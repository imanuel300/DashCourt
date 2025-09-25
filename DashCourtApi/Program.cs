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
bool useMockData = builder.Configuration.GetValue<bool>("AppSettings:UseMockData");
builder.Services.AddSingleton(new ExcelDataService(excelFilesPath, useMockData));

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
