# Listado de Radioaficionados Argentina

Dos scripts en Python:

- **`main.py`** descarga el padrón público de radioaficionados de ENACOM y lo guarda en Excel.
- **`predict_callsign.py`** usa ese padrón para listar las señales distintivas que todavía están libres para tu domicilio, y también puede exportar el resultado a Excel.

El archivo Excel que genera tiene cuatro hojas: **Resumen**, **Bloques**,
**Señales libres** (una fila por señal, para filtrar y ordenar) y **Licencias
locales** (si indicás una ciudad).

## Instalación

```bash
git clone https://github.com/marcosbtesh/listado-radio-aficionados-argentina.git
cd listado-radio-aficionados-argentina
```

Crear el entorno virtual e instalar las dependencias:

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

## Comandos

Descargar el padrón completo de ENACOM (genera `output/listado.xlsx`):

```bash
.venv/bin/python main.py
```

Ver las señales distintivas libres para tu provincia:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES"
```

Agregar tu ciudad, para ver también los radioaficionados ya licenciados ahí:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --city TIGRE
```

Consultar otra provincia (con o sin acentos):

```bash
.venv/bin/python predict_callsign.py --province CORDOBA
```

Filtrar por prefijo y numeral:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --prefix LU --numeral 3
```

Sufijo de dos letras (solo categoría SUPERIOR o ESPECIAL):

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --category SUPERIOR --suffix-length 2
```

Mostrar más resultados:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --top 30
```

Exportar el ranking completo a Excel (genera `output/senales_disponibles.xlsx`):

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --excel
```

Exportar a Excel eligiendo la ruta del archivo:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --excel mis_senales.xlsx
```

Salida en JSON, para procesar con otro programa:

```bash
.venv/bin/python predict_callsign.py --province "BUENOS AIRES" --json
```

Ver todas las opciones:

```bash
.venv/bin/python predict_callsign.py --help
```

## Cómo se forma una señal distintiva

Según el Reglamento General de Radioaficionados (Resolución ENACOM
[3635/2017](https://www.argentina.gob.ar/normativa/nacional/resoluci%C3%B3n-3635-2017-286986),
Capítulo VIII), la señal es **prefijo + número + sufijo**.

La ubicación geográfica está en las **primeras letras del sufijo**, no en el número
(punto 8.1). El número no indica ni provincia ni categoría.

| División | Letra | División | Letra |
|---|---|---|---|
| Ciudad Autónoma de Buenos Aires | A-B-C | Salta | O |
| Buenos Aires | D-E | San Juan | P |
| Santa Fe | F | San Luis | Q |
| Chaco | GA-GOZ | Catamarca | R |
| Formosa | GP-GZZ | La Rioja | S |
| Córdoba | H | Jujuy | T |
| Misiones | I | La Pampa | U |
| Entre Ríos | J | Río Negro | V |
| Tucumán | K | Chubut | W |
| Corrientes | L | Santa Cruz | XA-XOZ |
| Mendoza | M | Tierra del Fuego | XP-XZZ |
| Sgo. del Estero | N | Neuquén | Y |
| Antártida e Islas del Atlántico Sur | Z | | |

Otros puntos que aplica el script:

- **8.2.1** — el sufijo corresponde al domicilio de la estación fija (o al del DNI si no hay estación).
- **8.3** — el sufijo de dos letras solo se otorga, a pedido, a categoría SUPERIOR o ESPECIAL. El resto lleva tres letras.
- **4.10 / 8.4** — las señales caducadas se reasignan, así que los huecos en bloques viejos también son candidatos.

## Aviso

Los resultados son una **estimación**, no una predicción oficial. El criterio de
asignación es facultad exclusiva de ENACOM (punto 8.2). El script solo ordena las
señales libres según dónde viene asignando ENACOM últimamente.

## Problema conocido

Si `main.py` falla con un error de certificado (`CERTIFICATE_VERIFY_FAILED`), no es
un problema del script: el servidor de ENACOM envía solo su certificado final y
omite el intermedio de Sectigo. Los navegadores lo resuelven solos, Python no.
