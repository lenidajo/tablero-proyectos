# Tablero de Proyectos — CronoEstadística

Tablero de seguimiento de proyectos generado desde `CronoEstadistica!.xlsx`
(hoja "Crono") y publicado con GitHub Pages desde `/docs`.

## Archivos

- `template.html` — diseño del tablero (HTML/CSS/JS), con el placeholder
  `__DATA_JSON__` en lugar del arreglo de datos fijo.
- `actualizar_tablero.py` — lee el Excel fuente, genera `docs/index.html`
  a partir de `template.html` y hace `git add / commit / push`.
- `docs/index.html` — versión publicada del tablero (generada, no editar a mano).

## Actualizar manualmente

```
python actualizar_tablero.py
```

## Actualización automática

Tarea programada de Windows `TableroCronoEstadistica`, todos los días hábiles
a las 7:00 a. m., ejecuta este mismo script.
