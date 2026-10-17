# Extractor de datos — Dompavolei (Clupik)

Herramienta para extraer datos de la web del **Club Voleibol San Martiño
(Dompavolei)**, que funciona sobre la plataforma **Clupik**.

En lugar de raspar (scrapear) el HTML, el script consume directamente la **API
JSON pública de Clupik**, que no requiere iniciar sesión. Es más limpio, rápido
y estable.

## Qué extrae

De momento: **la lista de equipos del club**, exportada a un archivo Excel
(`.xlsx`).

Fuente de datos:

```
GET https://api.clupik.com/clubs/27574/teams
```

(`27574` es el ID del club. Los datos son públicos; cualquiera que visite la web
accede a ellos.)

## Instalación

Requiere Python 3.9 o superior.

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Genera equipos_dompavolei.xlsx con los equipos del club por defecto
python extraer_equipos.py

# Especificar otro club u otra ruta de salida
python extraer_equipos.py --club-id 27574 -o equipos.xlsx
```

El Excel resultante tiene dos hojas:

- **Equipos**: una fila por equipo (nombre, temporada, sección, visibilidad,
  escudo, IDs…), con encabezados fijos y filtros activados.
- **Info**: procedencia de los datos (club, URL de origen, nº de equipos y fecha
  de generación).

## Notas

- Solo se extraen **datos públicos**. Para datos privados (socios, pagos, etc.)
  la vía correcta es la exportación del panel de administración de Clupik, no
  este script.
- Si en el futuro quieres extraer **partidos/resultados** o **competiciones**,
  esos datos vienen de otro servicio de Clupik
  (`managergames.clupik.com`) y se pueden añadir al extractor.
