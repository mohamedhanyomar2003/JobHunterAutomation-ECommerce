var builder = WebApplication.CreateBuilder(args);


builder.Services.AddHostedService<GoogleSheetsWorker>();

 
builder.Services.AddControllers();

var app = builder.Build();

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.MapGet("/", () => "?? Job Hunter Bot is Running 24/7 on MonsterASP!");

app.Run();