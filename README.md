# Buscador de Clínicas - Detector de Oportunidades Web

Script de Python que busca clínicas en una zona determinada y detecta cuáles **no tienen página web** o tienen una **web muy básica**, para identificar oportunidades de negocio en desarrollo web.

## Qué hace

1. **Busca clínicas** en Google Places según la búsqueda que indiques (ej: "clínicas dentales en Madrid").
2. **Obtiene los datos** de cada clínica: nombre, teléfono, email, dirección y enlace web.
3. **Analiza la web** de cada clínica y la clasifica en:
   - **Sin web** — No tiene página web.
   - **Web inaccesible** — Tiene URL pero no responde.
   - **MUY BÁSICA / En construcción** — Web placeholder o en desarrollo.
   - **MUY BÁSICA** — Poco contenido y pocas secciones.
   - **BÁSICA** — Contenido o estructura limitada.
   - **ACEPTABLE** — Web funcional sin tecnología moderna.
   - **PROFESIONAL** — Web con tecnología moderna y buen contenido.
4. **Exporta todo a Excel** con formato profesional, filtros y resumen estadístico.

## Requisitos

- Python 3.10+
- Una **API key de Google** con la **Places API** habilitada.

### Obtener API key de Google Places

1. Ve a [Google Cloud Console](https://console.cloud.google.com/).
2. Crea un proyecto (o usa uno existente).
3. Habilita la **Places API** en "APIs y servicios".
4. Crea una API key en "Credenciales".

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Búsqueda básica
python buscar_clinicas.py --api-key TU_API_KEY --query "clínicas dentales en Madrid"

# Con más resultados y archivo de salida personalizado
python buscar_clinicas.py \
  --api-key TU_API_KEY \
  --query "clínicas estéticas en Barcelona" \
  --max-results 60 \
  --output clinicas_barcelona.xlsx

# Más hilos para mayor velocidad
python buscar_clinicas.py \
  --api-key TU_API_KEY \
  --query "fisioterapia en Valencia" \
  --workers 10
```

### Parámetros

| Parámetro | Corto | Descripción | Default |
|---|---|---|---|
| `--query` | `-q` | Texto de búsqueda | (requerido) |
| `--api-key` | `-k` | API key de Google Places | (requerido) |
| `--max-results` | `-m` | Máximo de resultados | 60 |
| `--output` | `-o` | Nombre del archivo Excel | `clinicas_prospecto.xlsx` |
| `--workers` | `-w` | Hilos paralelos para analizar webs | 5 |
| `--radius` | `-r` | Radio de búsqueda en metros | 5000 |

## Archivo Excel de salida

El archivo generado contiene dos hojas:

### Hoja "Clínicas - Prospectos"
| Columna | Descripción |
|---|---|
| Nombre | Nombre de la clínica |
| Teléfono | Número de contacto |
| Email | Emails encontrados en la web |
| Dirección | Dirección completa |
| Página Web | URL del sitio web |
| Google Maps | Enlace a Google Maps |
| Calidad Web | Clasificación de la web |
| Descripción | Detalle del análisis |
| Rating | Puntuación en Google |
| Nº Reseñas | Número de reseñas |
| ¿Oportunidad? | SÍ si no tiene web o es muy básica |

Las filas de oportunidad aparecen resaltadas en verde y ordenadas primero.

### Hoja "Resumen"
Estadísticas generales: total de clínicas, desglose por calidad de web y total de oportunidades.

## Ejemplos de búsquedas útiles

```
"clínicas dentales en Madrid"
"clínicas veterinarias en Sevilla"
"clínicas estéticas en Barcelona"
"fisioterapia en Málaga"
"podología en Bilbao"
"clínicas oftalmológicas en Zaragoza"
"centros de psicología en Granada"
"clínicas de fertilidad en España"
```
