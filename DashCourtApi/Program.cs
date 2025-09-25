using DashCourtApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Configure CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigin",
        builder => builder.WithOrigins("http://localhost:5500") // Replace with your frontend URL
                            .AllowAnyHeader()
                            .AllowAnyMethod());
});

// Register ExcelDataService as a singleton
string excelFilesPath = Path.Combine(Directory.GetCurrentDirectory(), "..", ".."); // Adjust this path if needed
builder.Services.AddSingleton(new ExcelDataService(excelFilesPath));

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
