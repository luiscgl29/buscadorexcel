using MiniExcelLibs;
using System.Globalization;
using System.Text;

var builder = WebApplication.CreateBuilder(args);

// Configuración para que funcione en LAN (WiFi)
var port = Environment.GetEnvironmentVariable("PORT") ?? "5000";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}"); 
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll",
        builder => builder.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader());
});

var app = builder.Build();
app.UseCors("AllowAll");

// --- RUTA DEL ARCHIVO EXCEL ---
//string rutaExcel = @"C:\BDEstudiantes\Estudiantes2026.xlsx";
string rutaExcel = Path.Combine(AppContext.BaseDirectory, "Data", "Estudiantes2026.xlsx");

// 1. FRONTEND (Página Web)
app.MapGet("/", () => Results.Content(@"
<!DOCTYPE html>
<html lang='es'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Buscador Escolar</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #eef2f7; }
        .container { max-width: 900px; margin: 0 auto; }
        
        input { width: 100%; padding: 15px; font-size: 18px; border-radius: 8px; border: 1px solid #ccc; box-shadow: 0 2px 5px rgba(0,0,0,0.1); box-sizing: border-box; }
        
        .card { background: white; margin-top: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.08); overflow: hidden; border-left: 6px solid #2980b9; }
        
        .card-header { background: #f8f9fa; padding: 15px 20px; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; align-items: center; }
        .nombre-alumno { font-size: 1.3em; font-weight: bold; color: #2c3e50; }
        .matricula { background: #2980b9; color: white; padding: 5px 10px; border-radius: 15px; font-size: 0.9em; font-weight: bold; }
        
        .card-body { padding: 20px; display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        @media (max-width: 600px) { .card-body { grid-template-columns: 1fr; } }

        .section-title { font-size: 0.85em; text-transform: uppercase; color: #7f8c8d; font-weight: bold; margin-bottom: 8px; border-bottom: 2px solid #ecf0f1; padding-bottom: 3px; }
        .data-row { margin-bottom: 6px; font-size: 0.95em; color: #34495e; }
        .label { font-weight: 600; color: #555; }
        .empty-msg { text-align:center; color:#7f8c8d; margin-top:20px; }
    </style>
</head>
<body>
    <div class='container'>
        <h2 style='text-align:center; color:#2c3e50;'>Búsqueda de Estudiantes</h2>
        <input type='text' id='buscador' placeholder='Escribir nombre, apellido o matrícula del estudiante...'>
        <div id='resultados'></div>
    </div>

    <script>
        let timeout = null;
        const input = document.getElementById('buscador');
        const resultadosDiv = document.getElementById('resultados');

        input.addEventListener('input', function (e) {
            clearTimeout(timeout);
            timeout = setTimeout(() => buscar(e.target.value), 300);
        });

        async function buscar(texto) {
            if(texto.length < 3) { resultadosDiv.innerHTML = ''; return; }
            
            try {
                const response = await fetch('/api/buscar?q=' + encodeURIComponent(texto));
                const data = await response.json();
                
                if (data.length === 0) {
                    resultadosDiv.innerHTML = '<p class=\'empty-msg\'>No se encontró el estudiante.</p>';
                    return;
                }

                const html = data.map(est => `
                    <div class='card'>
                        <div class='card-header'>
                            <div class='nombre-alumno'>${est.nombre}</div>
                            <span class='matricula'>${est.matricula}</span>
                        </div>
                        <div class='card-body'>
                            <div>
                                <div class='section-title'>📚 Datos del Estudiante</div>
                                <div class='data-row'><span class='label'>Grado:</span> ${est.grado} - ${est.seccion}</div>
                                <div class='data-row'><span class='label'>Nacimiento:</span> ${est.fechaNacEstudiante}</div>
                                <div class='data-row'><span class='label'>Dirección:</span> ${est.direccionEstudiante}</div>
                                <div class='data-row'><span class='label'>Plan de estudio:</span> ${est.planEstudio}</div>
                                <div class='data-row'><span class='label'>Celular:</span> ${est.celularAlumno}</div>
                            </div>
                            <div>
                                <div class='section-title'>👨‍👩‍👦 Familiares</div>
                                <div class='data-row'>
                                <div class='data-row'><span class='label'>Padre:</span> ${est.padreNombre}</div>
                                <div class='data-row'><span class='label'>Celular:</span> ${est.padreCel}</div>
                                <div class='data-row'><span class='label'>Fecha Nacimiento:</span> ${est.fechaNacPadre}</div>
                                <div class='data-row'><span class='label'>Profesion Padre:</span> ${est.profesionPadre}</div>
                                </div>
                                <div class='data-row' style='margin-top:10px; border-top:1px dashed #ccc; padding-top:5px;'>
                                    <div class='data-row'><span class='label'>Madre:</span> ${est.madreNombre}</div>
                                    <div class='data-row'><span class='label'>Celular:</span> ${est.madreCel}</div>
                                    <div class='data-row'><span class='label'>Fecha Nacimiento:</span> ${est.fechaNacMadre}</div>
                                    <div class='data-row'><span class='label'>Profesion Madre:</span> ${est.profesionMadre}</div>
                                </div>
                                <div class='data-row' style='margin-top:10px; border-top:1px dashed #ccc; padding-top:5px;'>
                                    <div class='data-row'><span class='label'>Encargado:</span> ${est.encargadoNombre}</div>
                                    <div class='data-row'><span class='label'>Celular:</span> ${est.encargadoCel}</div>
                                    <div class='data-row'><span class='label'>Fecha Nacimiento:</span> ${est.fechaNacEncargado}</div>
                                    <div class='data-row'><span class='label'>Profesion Madre:</span> ${est.profesionEncargado}</div>
                                </div>
                            </div>
                        </div>
                    </div>
                `).join('');
                resultadosDiv.innerHTML = html;
            } catch (err) { console.error(err); resultadosDiv.innerHTML = '<p class=\'empty-msg\'>Error de conexión.</p>'; }
        }
    </script>
</body>
</html>
", "text/html"));

// ============================================================
// ENDPOINT DE DIAGNÓSTICO — Visita /api/columnas para ver
// los nombres EXACTOS que MiniExcel lee de tu archivo.
// Puedes eliminarlo cuando todo funcione bien.
// ============================================================
app.MapGet("/api/columnas", () =>
{
    var filas = MiniExcel.Query(rutaExcel, useHeaderRow: true, startCell: "A4").Cast<IDictionary<string, object>>();
    var primera = filas.FirstOrDefault();
    if (primera == null) return Results.Ok("Sin datos");
    return Results.Ok(primera.Keys.ToList());
});

// 2. BACKEND (Lógica de Búsqueda)
app.MapGet("/api/buscar", (string q) =>
{
    if (string.IsNullOrWhiteSpace(q)) return Results.Ok(new List<object>());

    string busquedaNorm = NormalizarTexto(q);

    var filas = MiniExcel.Query(rutaExcel, useHeaderRow: true, startCell: "A4").Cast<IDictionary<string, object>>();

    var resultados = filas.Where(row =>
    {
        string nombre = ObtenerValorStr(row, "Apellido, Nombre");
        string matricula = ObtenerValorStr(row, "Matrícula");

        return NormalizarTexto(nombre).Contains(busquedaNorm) ||
               NormalizarTexto(matricula).Contains(busquedaNorm);
    })
    .Select(row => new
    {
        // Matrícula: quitar comilla simple inicial si existe (ej: '20260312 → 20260312)
        matricula = ObtenerValorStr(row, "Matrícula").Trim('\'', ' '),
        nombre = ObtenerValorStr(row, "Apellido, Nombre"),
        celularAlumno = ObtenerValorStr(row, "Celular alumno"),

        // ★ Fechas: se formatean a dd/MM/yyyy porque vienen como DateTime
        //-----------------------METER COMPROBACION SI ES MAYOR DE EDAD --------------------
        fechaNacEstudiante = FormatearFecha(row, "Fecha de nacimiento estudiante"), // Edad del estudiante

        //Pendiente borrar -----------------------------------------------------
        lugarNacEstudiante = ObtenerValorStr(row, "Lugar de nacimiento estudiante"),
        //----------------------------------------------------
        direccionEstudiante = ObtenerValorStr(row, "Dirección de casa estudiante"),

        grado = ObtenerValorStr(row, "Grado"),
        seccion = ObtenerValorStr(row, "Sección"),
        planEstudio = ObtenerValorStr(row, "Plan de estudio"),

        // Padre: nombre de col X, celular primero intenta AJ, luego DH
        padreNombre = ObtenerValorStr(row, "Nombre del padre"),
        fechaNacPadre = FormatearFecha(row, "Fecha de nacimiento del padre"), // Colocar edad exacta
        profesionPadre = ObtenerValorStr(row, "Profesión padre"),
        //lugarNacPadre = ObtenerValorStr(row, "Lugar de nacimiento de padre"),
        padreCel = ObtenerValorStr(row, "Celular de padre"),
        padreTel = ObtenerValorStr(row, "Teléfono de padre"),

        // Madre: nombre de col AO, celular primero intenta BA, luego DI
        madreNombre = ObtenerValorStr(row, "Nombre de la madre"),
        fechaNacMadre = FormatearFecha(row, "Fecha de nacimiento del padre"), // Colocar edad exacta
        profesionMadre = ObtenerValorStr(row, "Profesión madre"),
        //lugarNacMadre = ObtenerValorStr(row, "Lugar de nacimiento de madre"),
        madreCel = ObtenerValorStr(row, "Celular de la madre"),
        madreTel = ObtenerValorStr(row, "Teléfono de la madre"),
        

        // Encargado: col DJ para nombre, col DO para teléfono

        encargadoNombre = ObtenerValorPreferente(row, "Nombre encargado", "Nombre de Encargado"),
        fechaNacEncargado = FormatearFecha(row, "Fecha de nacimiento encargado"),
        encargadoCel = ObtenerValorStr(row, "Celular encargado"),
        encargadoTel = ObtenerValorStr(row, "Teléfono casa encargado"),02
        profesionEncargado = ObtenerValorStr(row, "Profesión encargado"),

    })
    .Take(15)
    .ToList();

    return Results.Ok(resultados);
});

app.Run();

// =====================================================
// FUNCIONES DE AYUDA
// =====================================================

/// <summary>
/// Obtiene el valor de una columna como STRING, sin importar si
/// el dato original es int, double, DateTime, etc.
/// Busca la columna ignorando tildes, espacios y mayúsculas.
/// </summary>
static string ObtenerValorStr(IDictionary<string, object> fila, string columnaBuscada)
{
    if (fila == null) return "";
    var llaveReal = fila.Keys.FirstOrDefault(k =>
        NormalizarParaColumna(k) == NormalizarParaColumna(columnaBuscada));

    if (llaveReal == null) return "---";

    var valor = fila[llaveReal];
    if (valor == null) return "";

    // Convertir todo a string limpio
    return valor.ToString() ?? "";
}

/// <summary>
/// Formatea columnas que contienen fechas (DateTime) a "dd/MM/yyyy".
/// Si no es DateTime, devuelve el texto tal cual.
/// </summary>
static string FormatearFecha(IDictionary<string, object> fila, string columnaBuscada)
{
    if (fila == null) return "";
    var llaveReal = fila.Keys.FirstOrDefault(k =>
        NormalizarParaColumna(k) == NormalizarParaColumna(columnaBuscada));

    if (llaveReal == null) return "---";

    var valor = fila[llaveReal];
    if (valor == null) return "";

    // Si es DateTime, formatear bonito
    if (valor is DateTime dt)
    {
        var hoy = DateTime.Today;
        int edad = hoy.Year - dt.Year;

        if(dt.Date > hoy.AddYears(-edad))
        {
            edad--;
        }
        return $"{dt:dd/MM/yyyy} / {edad} años";
    }

    return valor.ToString() ?? "";
}

/// <summary>
/// Busca en col1 primero; si está vacía o es "---", busca en col2.
/// Útil para columnas duplicadas como "Celular de padre" / "Celular Padre".
/// </summary>
static string ObtenerValorPreferente(IDictionary<string, object> fila, string col1, string col2)
{
    string valor1 = ObtenerValorStr(fila, col1);
    if (!string.IsNullOrWhiteSpace(valor1) && valor1 != "---") return valor1;
    return ObtenerValorStr(fila, col2);
}

/// <summary>
/// Normaliza nombres de columnas: quita tildes, espacios, saltos de línea, pasa a minúsculas.
/// Así "Apellido,\n Nombre" == "apellido,nombre"
/// </summary>
static string NormalizarParaColumna(string texto)
{
    if (string.IsNullOrEmpty(texto)) return "";
    var sb = new StringBuilder();
    foreach (var c in texto.Normalize(NormalizationForm.FormD))
    {
        if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
            sb.Append(c);
    }
    return sb.ToString().Normalize(NormalizationForm.FormC)
        .ToLowerInvariant()
        .Replace(" ", "")
        .Replace("\n", "")
        .Replace("\r", "");
}

/// <summary>
/// Normaliza texto de búsqueda: quita tildes y pasa a minúsculas.
/// Así "López" == "lopez" y "GARCIA" == "garcia".
/// </summary>
static string NormalizarTexto(string texto)
{
    if (string.IsNullOrEmpty(texto)) return "";
    var sb = new StringBuilder();
    foreach (var c in texto.Normalize(NormalizationForm.FormD))
    {
        if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
            sb.Append(c);
    }
    return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
}