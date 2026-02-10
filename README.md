Sistema Comparador de Compras Inteligente con IA
https://img.shields.io/badge/PowerShell-5.1+-blue.svg?style=for-the-badge&logo=powershell
https://img.shields.io/badge/Windows-7%7C8%7C10%7C11-0078D6?style=for-the-badge&logo=windows
https://img.shields.io/badge/Excel-Macros%2520Enabled-217346?style=for-the-badge&logo=microsoftexcel
https://img.shields.io/badge/Estado-Versi%C3%B3n%25203.5%2520(Estable)-brightgreen?style=for-the-badge
https://img.shields.io/badge/Licencia-MIT-yellow?style=for-the-badge

📋 Descripción
Sistema Comparador de Compras Inteligente con IA es una solución integral que evoluciona desde un sistema basado en Excel con macros hacia una aplicación completa con inteligencia artificial. Diseñado inicialmente para uso personal pero con arquitectura multi-usuario desde el inicio, permite comparar precios, optimizar rutas de compra y analizar tendencias de consumo con algoritmos de recomendación personalizados.

El sistema combina la accesibilidad de Excel con la potencia de scripts de automatización (PowerShell, Batch) y está preparado para escalar hacia módulos de machine learning y una futura aplicación web/móvil independiente.

🏗️ Arquitectura del Sistema
text
Comparador_Compras_IA/
├── 📁 tools/                         # Scripts de instalación y automatización
│   ├── 🔧 crear_sistema.bat         # Instalador principal (v3.5)
│   ├── 📝 crear_excel.ps1          # Generador de estructura Excel (v4.0)
│   ├── 📊 cargar_datos.ps1         # Carga de datos de ejemplo
│   ├── ⚙️ configurar_sistema.ps1   # Configuración avanzada (v3.5)
│   └── 📄 README_SCRIPTS.txt       # Documentación específica de scripts
│
├── 📁 src/                          # Código fuente principal
│   ├── 📁 vba_modules/             # Módulos VBA para Excel
│   ├── 📁 powershell_libs/         # Librerías PowerShell reutilizables
│   └── 📁 future_ia/               # Estructura para futuros módulos IA
│
├── 📁 docs/                         # Documentación completa
│   ├── 📄 ARCHITECTURE.md          # Arquitectura técnica detallada
│   ├── 📄 API_REFERENCE.md         # Referencia de APIs (futuro)
│   └── 📄 USER_GUIDE.md            # Manual de usuario completo
│
├── 📁 samples/                      # Ejemplos y datos de prueba
│   ├── 📁 config_templates/        # Plantillas de configuración
│   └── 📁 sample_data/             # Datos de ejemplo para pruebas
│
├── 📄 .gitignore                   # Archivos excluidos de Git
├── 📄 LICENSE                      # Licencia MIT
├── 📄 CHANGELOG.md                 # Historial de cambios
└── 📄 README.md                    # Este archivo
✨ Características Principales
✅ Instalación y Configuración Robusta
Instalador en 8 fases con verificación exhaustiva del sistema

Compatibilidad total con Windows 7, 8, 10 y 11 (32/64/ARM)

Backup automático antes de reinstalación

Sistema de logs completo y organizado por categorías

📊 Gestión de Datos Avanzada
10 tablas interrelacionadas en Excel con relaciones complejas

Sistema de backup 3-2-1 (diario, semanal, mensual)

Importación/exportación en múltiples formatos (CSV, JSON, XML, PDF)

Validación de datos integrada y estructura modular

🧠 Funcionalidades Inteligentes
Algoritmo de comparación con ponderación personalizable

Optimización de rutas considerando distancia, tiempo y coste

Sistema de recomendaciones basado en historial de compras

Preparado para IA con arquitectura escalable para machine learning

🔒 Seguridad y Estabilidad
Manejo de errores mejorado en todos los componentes

Verificación de permisos y compatibilidad del sistema

Configuración jerárquica (sistema → usuario → sesión)

Recuperación automática en caso de fallos críticos

🚀 Instalación Rápida
Prerrequisitos
Sistema Operativo: Windows 7 SP1 o superior

Microsoft Excel: 2013 o superior (para todas las funciones)

PowerShell: Versión 3.0 o superior (incluido en Windows)

Espacio en disco: Mínimo 500 MB recomendados

Pasos de Instalación
Clonar el repositorio

bash
git clone https://github.com/MarioF7/Comparador_Compras_IA.git
cd Comparador_Compras_IA
Ejecutar el instalador principal

powershell
# Navegar a la carpeta de herramientas
cd tools

# Ejecutar como administrador (recomendado)
.\crear_sistema.bat
Seguir el proceso guiado
El instalador ejecutará 8 fases automáticas:

text
FASE 1: Verificación del sistema (OS, PowerShell, .NET, espacio)
FASE 2: Preparación del entorno (backup de instalación anterior)
FASE 3: Creación de estructura (15 carpetas principales, 58 subcarpetas)
FASE 4: Ejecución de scripts de configuración
FASE 5: Creación de archivos de configuración
FASE 6: Creación de accesos directos
FASE 7: Verificación final del sistema
FASE 8: Resumen y finalización
Iniciar el sistema

Buscar el acceso directo "Comparador Compras IA" en el escritorio

Abrir Comparador_Compras_IA_Completo.xlsm

Habilitar macros cuando Excel lo solicite

Completar datos iniciales en la hoja USUARIOS

📖 Uso del Sistema
Primeros Pasos
Configuración inicial: Complete su perfil en la hoja USUARIOS

Añadir tiendas: Registre supermercados locales en la hoja TIENDAS

Cargar productos: Añada productos frecuentes en la hoja PRODUCTOS

Actualizar precios: Ingrese precios actuales en la hoja PRECIOS

Funciones Principales
excel
1. COMPARACIÓN SIMPLE
   - Seleccionar productos a comparar
   - Elegir radio de búsqueda (km)
   - Generar reporte de mejores precios

2. ANÁLISIS DE TENDENCIAS
   - Ver historial de precios por producto
   - Identificar patrones estacionales
   - Recibir alertas de cambios de precio

3. OPTIMIZACIÓN DE RUTAS
   - Calcular ruta más eficiente para múltiples compras
   - Considerar horarios de tiendas y tráfico
   - Optimizar coste total (productos + desplazamiento)

4. REPORTES AVANZADOS
   - Generar dashboards interactivos
   - Exportar a PDF/Excel para compartir
   - Análisis estadístico de hábitos de compra
Scripts de Utilidad Incluidos
powershell
# Backup automático programado
.\tools\backup_automatico.ps1 -ProjectPath "C:\Tu_Ruta"

# Diagnóstico del sistema
.\tools\verificar_sistema.ps1 -DetailedReport

# Limpieza de caché y temporales
.\tools\limpiar_cache.ps1 -ProjectPath "C:\Tu_Ruta"
🔧 Tecnologías Utilizadas
Componente	Tecnología	Versión	Propósito
Base de Datos	Excel + VBA	2016+	Almacenamiento estructurado y cálculos
Automatización	PowerShell	5.1+	Scripts de instalación y mantenimiento
Interfaz	Excel Forms	-	Formularios de usuario y controles
Configuración	JSON/XML	-	Configuración jerárquica del sistema
Backup	ZIP + Robocopy	-	Sistema de respaldo multi-nivel
Logging	Texto estructurado	-	Registro de eventos y auditoría
📈 Estado del Proyecto
✅ Completado (Versión 3.5)
Instalador robusto con 8 fases verificadas

Estructura completa de 15 carpetas principales

Sistema de configuración jerárquico (JSON/XML)

Scripts de utilidad (backup, verificación, limpieza)

Documentación técnica completa

Compatibilidad multiplataforma Windows

🔄 En Desarrollo
Macros VBA para funcionalidad básica

Dashboard interactivo en Excel

Sistema de importación/exportación mejorado

Fórmulas de cálculo optimizadas

⏳ Planeado (Roadmap)
Fase 3: Web scraping automático de precios

Fase 4: APIs externas (Google Maps, supermercados)

Fase 5: Sistema de recomendación con machine learning

Fase 6: Aplicación web/móvil independiente

🐛 Solución de Problemas Comunes
Problema: Excel abre como "solo lectura"
Solución:

powershell
# Ejecutar desde PowerShell como administrador
Unblock-File -Path "Comparador_Compras_IA_Completo.xlsm"
# O hacer clic derecho → Propiedades → Desbloquear
Problema: Error de ejecución de PowerShell
Solución:

powershell
# Ejecutar en PowerShell como administrador
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
Problema: Macros no se ejecutan
Solución:

Ir a Archivo → Opciones → Centro de confianza

Configuración del Centro de confianza → Configuración de macros

Seleccionar "Habilitar todas las macros"

🤝 Cómo Contribuir
Las contribuciones son bienvenidas y apreciadas. Por favor:

Reportar bugs a través de Issues

Sugerir mejoras con la plantilla de feature request

Enviar Pull Requests para correcciones o nuevas funcionalidades

Guía de contribución:
bash
# 1. Hacer fork del repositorio
# 2. Crear una rama para tu funcionalidad
git checkout -b feature/nueva-funcionalidad

# 3. Commit de cambios con mensajes descriptivos
git commit -m "FEAT: Añade comparación multi-producto"

# 4. Push a tu fork
git push origin feature/nueva-funcionalidad

# 5. Abrir Pull Request con descripción detallada
📄 Licencia
Este proyecto está licenciado bajo la Licencia MIT - ver el archivo LICENSE para detalles completos.

Resumen de licencia:
✅ Uso comercial permitido

✅ Modificación y distribución permitidas

✅ Uso privado permitido

✅ Incluir licencia y aviso de copyright

✅ Sin garantía

✅ No responsabilidad del autor

🙏 Agradecimientos
Equipo de desarrollo: Por la arquitectura modular y escalable

Comunidad open-source: Por las herramientas y librerías utilizadas

Testers beta: Por su invaluable feedback y reporte de bugs

📞 Soporte y Contacto
Reportar problemas: Issues del repositorio

Documentación completa: Wiki del proyecto

Email de contacto: [Incluir si es relevante]

¿Listo para optimizar tus compras? ⚡

text
¡Instala, configura y comienza a ahorrar hoy mismo!
El sistema más completo para la comparación inteligente de precios.
Última actualización: Febrero 2024 | Versión del sistema: 3.5.0