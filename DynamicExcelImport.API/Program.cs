using DynamicExcelImport.API.Services; //servisin klasör adýný ekledik.
var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddScoped<DatabaseService>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddScoped<DynamicExcelImport.API.Services.IExcelProcessingService, DynamicExcelImport.API.Services.ExcelProcessingService>();
builder.Services.AddScoped<DynamicExcelImport.API.Services.IReportManagementService, DynamicExcelImport.API.Services.ReportManagementService>();
builder.Services.AddScoped<DynamicExcelImport.API.Services.IDynamicDataService, DynamicExcelImport.API.Services.DynamicDataService>();
builder.Services.AddSwaggerGen();


var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
