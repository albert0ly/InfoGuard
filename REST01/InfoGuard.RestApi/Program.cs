using System.Numerics;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddOpenApi();

var allowedAudiences = builder.Configuration.GetSection("Entra:AllowedAudiences").Get<string[]>()
    ?? Array.Empty<string>();
var corsOrigins = builder.Configuration.GetSection("Cors:AllowedOrigins").Get<string[]>()
    ?? Array.Empty<string>();

builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = builder.Configuration["Entra:Authority"];
        options.TokenValidationParameters.ValidateAudience = true;
        options.TokenValidationParameters.ValidAudiences = allowedAudiences;
        options.TokenValidationParameters.ValidateIssuer = false;
    });

builder.Services.AddAuthorization();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AddinClientPolicy", policy =>
    {
        policy.WithOrigins(corsOrigins)
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseHttpsRedirection();
app.UseCors("AddinClientPolicy");
app.UseAuthentication();
app.UseAuthorization();

app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

app.MapPost("/api/fibonacci", [Authorize] (FibonacciRequest request) =>
{
    if (request.Number < 0)
    {
        return Results.BadRequest(new { error = "Number must be non-negative." });
    }

    var fibonacci = CalculateFibonacci(request.Number);
    return Results.Ok(new FibonacciResponse(request.Number, fibonacci.ToString()));
});

app.Run();

static BigInteger CalculateFibonacci(int n)
{
    if (n == 0)
    {
        return BigInteger.Zero;
    }

    if (n == 1)
    {
        return BigInteger.One;
    }

    var prev = BigInteger.Zero;
    var current = BigInteger.One;

    for (var i = 2; i <= n; i++)
    {
        var next = prev + current;
        prev = current;
        current = next;
    }

    return current;
}

public record FibonacciRequest(int Number);

public record FibonacciResponse(int Number, string Fibonacci);
