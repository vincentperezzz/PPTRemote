using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace PptRemote;

internal static class AppServer
{
    private static readonly byte[] Pixel = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static WebApplication Start(int port, PowerPointService ppt, ClientHub hub)
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ContentRootPath = AppContext.BaseDirectory,
            ApplicationName = "PptRemote"
        });
        builder.Logging.ClearProviders();
        builder.WebHost.UseUrls($"http://0.0.0.0:{port}");
        var app = builder.Build();
        app.Use(async (ctx, next) =>
        {
            ctx.Response.Headers.CacheControl = "no-store";
            await next();
            hub.Touch(ctx.Connection.RemoteIpAddress?.ToString(), ctx.Request.Path.Value ?? "");
        });

        app.MapGet("/", () =>
        {
            var html = WebFiles.Text("index.html");
            var json = JsonSerializer.Serialize(ppt.Snapshot(), JsonOpts);
            html = html.Replace("null;/*BOOT*/", json + ";");
            return Results.Content(html, "text/html; charset=utf-8");
        });
        app.MapGet("/app.js", () => Results.Text(WebFiles.Text("app.js"), "text/javascript; charset=utf-8"));
        app.MapGet("/styles.css", () => Results.Text(WebFiles.Text("styles.css"), "text/css; charset=utf-8"));
        app.MapGet("/live.js", () =>
        {
            var json = JsonSerializer.Serialize(ppt.Snapshot(), JsonOpts);
            return Results.Text("window.__BOOT=" + json + ";", "application/javascript; charset=utf-8");
        });
        app.MapGet("/viewers", () => Results.Json(hub.Active(), JsonOpts));
        app.MapGet("/go/next", () => Click(() => ppt.Next()));
        app.MapGet("/go/prev", () => Click(() => ppt.Prev()));
        app.MapGet("/go/first", () => Click(() => ppt.First()));
        app.MapGet("/go/last", () => Click(() => ppt.Last()));
        app.MapGet("/go/start", () => Click(() => ppt.StartShow()));
        app.MapGet("/go/end", () => Click(() => ppt.EndShow()));
        app.MapGet("/go/black", () => Click(() => ppt.Black()));
        app.MapGet("/go/white", () => Click(() => ppt.White()));
        app.MapGet("/go/goto", (int n) => Click(() => ppt.Goto(n)));
        app.MapGet("/thumbs/{index:int}.png", (int index) =>
        {
            var path = Path.Combine(ppt.ThumbDir, index + ".png");
            if (!File.Exists(path))
            {
                return Results.NotFound();
            }

            return Results.File(path, "image/png");
        });
        app.MapGet("/ahead.png", () =>
        {
            var path = ppt.AheadFile;
            if (!File.Exists(path))
            {
                return Results.NotFound();
            }

            return Results.File(path, "image/png");
        });

        _ = app.RunAsync();
        return app;
    }

    private static IResult Click(Func<PresenterState> run)
    {
        try
        {
            run();
        }
        catch
        {
        }

        return Results.File(Pixel, "image/gif");
    }
}
