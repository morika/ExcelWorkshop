using ExcelWorkshop;
using PackageRepository.Components.Spreadsheet;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

var summaries = new[]
{
    "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
};

app.MapPost("/ReadFile", (IFormFile file) =>
{
    using MemoryStream memoryStream = new();
    file.CopyTo(memoryStream);
    var spreadsheet = new Spreadsheet<WeatherForecast>();
    var fileContent = spreadsheet.Read(memoryStream, "Sheet1", 1, 2);
    List<WeatherForecast> cast = fileContent;
    return cast;
}).DisableAntiforgery();

app.MapPost("/WriteFile", (List<WeatherForecast> request) =>
{
    var spreadsheet = new Spreadsheet<WeatherForecast>();
    spreadsheet.Write(request, "Sheet1", 1);
}).DisableAntiforgery();

app.Run();