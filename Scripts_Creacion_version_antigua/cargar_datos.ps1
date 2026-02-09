# Script para cargar datos iniciales en el sistema
Write-Host "Cargando datos iniciales..." -ForegroundColor Yellow

# Crear archivo de usuarios si no existe
$usuariosPath = "usuarios.csv"
if (-not (Test-Path $usuariosPath)) {
    @"
ID_Usuario,Nombre,Apodo,Direccion,Coordenadas,Radio_Busqueda_km,Prioridad_Compra,Restricciones,Presupuesto_Mensual_€,Email,Telefono
U001,Tu Nombre,Yo,Tu Dirección,40.4168,-3.7038,10,Precio-Calidad,Ninguna,600,tuemail@ejemplo.com,600123456
U002,Usuario Familiar,Familia,Calle Familiar 123,40.4200,-3.7000,15,Calidad-Precio,Sin gluten,800,familia@ejemplo.com,600654321
"@ | Out-File -FilePath $usuariosPath -Encoding UTF8
    Write-Host "✓ Archivo de usuarios creado" -ForegroundColor Green
}

# Crear archivo de productos básicos
$productosPath = "productos.csv"
if (-not (Test-Path $productosPath)) {
    @"
ID_Producto,Nombre,Categoria,Marca,Peso_Volumen,Unidad,Precio_Medio_€,Nutriscore,Ecologico
P001,Leche Entera,Lácteos,Pascual,1,L,0.95,B,No
P002,Huevos M,Lácteos,Camperos,12,ud,2.50,A,Sí
P003,Pan Integral,Panadería,Bimbo,400,g,1.20,B,No
P004,Plátanos,Frutas,,1,kg,1.80,A,Sí
P005,Tomates,Verduras,,1,kg,1.50,A,Sí
P006,Pollo,Carnes,Campofrío,1,kg,6.50,B,No
P007,Arroz,Legumbres,Brillante,1,kg,1.10,A,No
P008,Aceite Oliva,Aceites,Carbonell,1,L,6.50,C,Sí
P009,Café,Bebidas,Marcilla,250,g,4.50,C,No
P010,Yogur,Lácteos,Danone,125,ml,0.35,A,Sí
"@ | Out-File -FilePath $productosPath -Encoding UTF8
    Write-Host "✓ Archivo de productos creado" -ForegroundColor Green
}

# Crear archivo de tiendas
$tiendasPath = "tiendas.csv"
if (-not (Test-Path $tiendasPath)) {
    @"
ID_Tienda,Nombre,Cadena,Direccion,Distancia_km,Tiempo_min,Valoracion
T001,Mercadona,Mercadona,Calle Principal 123,2.5,15,4.2
T002,Carrefour,Carrefour,Avenida Central 456,3.8,22,4.0
T003,DIA,DIA,Plaza Pequeña 789,1.2,8,3.8
T004,Lidl,Lidl,Calle Secundaria 101,4.5,25,4.1
T005,Aldi,Aldi,Calle Nueva 202,5.2,28,4.3
"@ | Out-File -FilePath $tiendasPath -Encoding UTF8
    Write-Host "✓ Archivo de tiendas creado" -ForegroundColor Green
}

# Crear carpeta para tickets si no existe
if (-not (Test-Path "Tickets")) {
    New-Item -ItemType Directory -Path "Tickets" | Out-Null
    Write-Host "✓ Carpeta Tickets creada" -ForegroundColor Green
}

# Crear archivo de configuración
$configPath = "configuracion.txt"
if (-not (Test-Path $configPath)) {
    @"
[ConfiguracionSistema]
Version=1.0
FechaInstalacion=$(Get-Date -Format "yyyy-MM-dd")
UsuarioPrincipal=U001
RadioBusquedaDefault=10
Idioma=ES
Moneda=EUR
IVA=21

[API_Keys]
GoogleMaps=
OpenFoodFacts=
"@ | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "✓ Archivo de configuración creado" -ForegroundColor Green
}

Write-Host ""
Write-Host "✅ Datos iniciales cargados exitosamente" -ForegroundColor Green
Write-Host ""
Write-Host "Archivos creados:" -ForegroundColor Cyan
Write-Host "- usuarios.csv" -ForegroundColor White
Write-Host "- productos.csv" -ForegroundColor White
Write-Host "- tiendas.csv" -ForegroundColor White
Write-Host "- configuracion.txt" -ForegroundColor White
Write-Host "- Carpeta Tickets/" -ForegroundColor White

# Pausa al final
Write-Host ""
Write-Host "Presiona Enter para continuar..."
$null = Read-Host