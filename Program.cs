using PaySpaceWaitingEvents.API.Services;
using Microsoft.OData.Client;
using Microsoft.EntityFrameworkCore;
using PaySpaceWaitingEvents.API.Data;
using Payspace.Rest.API;


var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();

builder.Services.AddScoped<IWaitingEventsService, WaitingEventsService>();
builder.Services.AddScoped<IPayElementMappingService, PayElementMappingService>();
builder.Services.AddScoped<IUploadHistoryService, UploadHistoryService>();
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddScoped<DbContext, ApplicationDbContext>();
builder.Services.AddScoped<IPaySpaceApiService, PaySpaceApiService>();

builder.Services.AddSingleton<Payspace.Rest.API.Api>(sp =>
{
    var configuration = sp.GetRequiredService<IConfiguration>();
    var logger = sp.GetRequiredService<ILogger<Program>>();

    try
    {
        var clientId = configuration["PaySpace:ClientId"] ?? "";
        var secret = configuration["PaySpace:Secret"] ?? "";

        if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(secret))
        {
            logger.LogError("PaySpace API client ID or secret is missing");
            throw new InvalidOperationException("PaySpace API client ID or secret is missing");
        }

        logger.LogInformation("Initializing PaySpace API client");
        logger.LogInformation($"Using client ID: {clientId}");

        Payspace.Rest.API.TokenProvider.SetCredentials(clientId, secret);

        var apiClient = new Payspace.Rest.API.Api(clientId, secret, "", "");
        logger.LogInformation("PaySpace API client initialized successfully");
        return apiClient;
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Failed to initialize PaySpace API client");
        throw;
    }
});


builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", builder =>
    {
        builder.AllowAnyOrigin()
               .AllowAnyMethod()
               .AllowAnyHeader();
    });
});




var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseDefaultFiles();

app.UseStaticFiles();

app.UseCors("AllowAll");

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();