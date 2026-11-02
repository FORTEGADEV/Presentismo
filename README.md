# Presentismo

Lee el Excel mensual de control horario (una hoja por local, un bloque por
persona) y devuelve quién cobra presentismo y quién no.

## Cómo decide

Por persona, sumando todas las hojas donde aparezca:

- **Minutos tarde**: suma de la columna `PRESENTISMO` del bloque (se carga a mano en el Excel).
- **Faltas**: cada celda del bloque cuyo texto contiene "falta" (`Falta / Muela`, `falta con aviso`, ...).
- Cobra presentismo si `minutos_tarde <= --max-late-minutes` y `faltas que cuentan <= --max-faltas`.
- La hoja `Informe` se ignora.
- Si el bloque tiene el resumen manual (`PRESENTISMO: SI/NO`), se compara y se avisa si no coincide.

## Uso

```bash
pip install openpyxl
python presentismo.py Control_Horario_JULIO_2026.xlsx
python presentismo.py archivo.xlsx --max-late-minutes 15 --excusar "aviso|certific" --csv salida.csv
```

| Opción | Qué hace | Default |
|---|---|---|
| `--max-late-minutes N` | minutos tarde tolerados en el mes | 0 |
| `--max-faltas N` | faltas toleradas en el mes | 0 |
| `--excusar REGEX` | no cuenta las faltas cuyo texto coincide | ninguna |
| `--csv RUTA` | guarda el resultado en CSV | - |

## Tests

```bash
pip install pytest
pytest
```

## Límites conocidos

- No calcula las llegadas tarde desde los fichajes (los horarios varían por persona y día); usa los minutos que ya están anotados.
- No detecta días sin fichar si nadie anotó la falta en el Excel.
