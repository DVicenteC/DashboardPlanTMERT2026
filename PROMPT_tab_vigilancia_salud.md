# Prompt — Tab "Vigilancia de la Salud" en el dashboard TMERT 2026

Trabajaremos en: `C:\EspecialidadesTecnicas\_Proyectos_Git\dashboards\TMERT-Plan2026\TMERTDashboardProgramacion2026_EP.py`
Ambiente: `C:\ProgramData\miniconda3\envs\dash\python.exe` | NO modifiques ningún otro archivo.

---

## OBJETIVO

Agregar un **tab nuevo y separado** de Vigilancia de la Salud (VS) del Protocolo TMERT.

La pregunta que el tab tiene que contestar de un vistazo:
**¿qué centros de trabajo tienen condición Crítica (C) o Media/No Crítica (M) — o sea, requieren
vigilancia de la salud — y están pendientes?**

"Pendiente" tiene dos sabores distintos y hay que poder distinguirlos:
1. **No ha ingresado a vigilancia**: tiene condición C o M pero NO tiene nómina VS.
2. **Ingresó pero está incompleto**: tiene nómina VS y `evaluados < deben ingresar`.

Ojo con el caso 2: es el más común (31 de los 40 CT C/M con nómina en el plan) y suele venir con
`Fecha última Vigilancia de Salud` vacía, porque a nadie del CT le han tomado el examen aún. Un CT
con `deben ingresar = 30` y `evaluados = 0` es un pendiente severo, no un dato faltante.

---

## 1. LOS DATOS YA ESTÁN — NO LOS RECALCULES

`seguimientoTMERT_acelerada + Anexo 7.2V2.py` ya calcula la vigilancia de la salud y la sube a la
hoja de seguimiento de Google Sheets. El dashboard la lee en `cargar_datos_seguimiento_tmert()`
(→ `df_seg_raw` → `df_seg`). **No leas Provisa ni ninguna otra fuente desde el dashboard.**

Columnas nuevas ya disponibles en `df_seg` (nombres EXACTOS, con tildes y `°`; esta hoja NO pasa por
`normalizar_columnas_tmert`, sólo por `.str.strip()`):

| Columna | Qué es |
|---|---|
| `Condición o Nivel de Riesgo (C,M,A)` | Letra única **C / M / A**. Es la condición de la ÚLTIMA Identificación Avanzada; si el mismo día hubo varias, ya viene resuelta a la MÁS SEVERA (C>M>A). **Usa ésta.** |
| `Condición Identificación Avanzada (istprod)` | La glosa larga de lo mismo. Ya se usa en el tab "Estado Seguimiento"; sirve sólo para mostrar. |
| `N° de hombres que deben ingresar a vigilancia de salud 2026` | Nómina VS, RUT distintos. |
| `N° de mujeres que deben ingresar a vigilancia de salud 2026` | ídem. |
| `N° de hombres evaluados 2026` | De la nómina, los que tienen `Fecha_Evaluacion` informada. |
| `N° de mujeres evaluadas 2026` | ídem. |
| `Fecha última Vigilancia de Salud 2026` | Último **examen** de vigilancia del CT. Sólo 116 de los 172 CT con nómina la tienen: **si está vacía es porque a nadie de ese CT le han tomado el examen todavía**, no porque falte el dato. |
| `Fecha evaluación Vigilancia de Salud 2026` | Última **evaluación médica** del CT (102 CT). NO es la misma fecha que la anterior: difieren en 84 CT. |
| `Fecha Identificación Avanzada (real)` | Fecha de la identificación que originó la condición. |

**Trampas de tipo que sí importan:**

- La hoja se lee con `dtype=str`. Las 4 columnas de conteo llegan como texto (`"12"` o `""`).
  Conviértelas con `pd.to_numeric(..., errors='coerce')` y **deja los NaN como NaN, NO los rellenes
  con 0**: "sin nómina VS" (NaN) y "nómina de 0 personas" son cosas distintas, y justamente esa
  diferencia es lo que separa el pendiente tipo 1 del resto.
- Las columnas con `Fecha` en el nombre ya vienen parseadas a datetime por `cargar_datos_seguimiento_tmert()`.
- `Condición o Nivel de Riesgo (C,M,A)` viene como texto; ojo con los vacíos (`''`, `'nan'`, `'None'`).
- Cuidado con la propagación de nulos al filtrar: `~col.isin(['C','M'])` NO captura los nulos.
  Si quieres "los que no son C ni M", hazlo explícito (`.fillna('')` antes o `| col.isna()`).

**Alcance temporal — ojo con esto:** las 4 columnas de conteo y las dos fechas de VS cubren la
**ventana 2025-2026**, aunque el nombre de la columna diga "2026" (los nombres se mantuvieron para no
romper lo que ya consume la hoja). El conteo es de **RUT distintos en toda la ventana**: 1.064
trabajadores aparecen en los dos años y suman una sola vez, así que **NO sumes año a año** ni
interpretes estos números como anuales. Las fechas de vigilancia ambiental (identificación,
evaluación, prescripción) no tienen filtro de año en absoluto. No los compares como si fueran lo mismo.

---

## 2. QUÉ DEBE MOSTRAR EL TAB

Ordenado de "alerta" a "detalle". Sé sobrio: métricas + una tabla accionable valen más que muchos gráficos.

**a) Semáforo de pendientes (arriba, `st.metric`)**
- CT con condición C o M (desglosado C / M).
- De ellos: **sin ingresar a vigilancia** (sin nómina VS) — este es el número que el equipo persigue.
- De ellos: **con vigilancia incompleta** (`evaluados < deben`).
- De ellos: completos (`evaluados >= deben`).
- Trabajadores: total que deben ingresar vs total evaluados, con el % de cobertura.

**b) Tabla accionable de pendientes** — el corazón del tab.
Filas: los CT con condición C o M que no están completos. Columnas sugeridas:
`Región, Ergonomo, Nombre Empleador, ID-CT, Nombre CT, Comuna CT, Estado Centro de Trabajo,
Condición (C/M), Fecha Ident. Avanzada, Días desde Ident. Avanzada, Estado VS
(Sin ingresar / Incompleta / Completa), Deben H, Deben M, Evaluados H, Evaluados M, Brecha,
Fecha última VS (examen), Fecha últ. evaluación VS`.
- Ordenar por criticidad y antigüedad: **primero C, luego M; dentro de cada grupo, mayor
  "días desde Ident. Avanzada" arriba** (los sin fecha primero o al final, tú decides, pero déjalo dicho).
- Resaltar en rojo suave (`#fdecea`, mismo patrón que el bloque de Críticas ya existente) las filas
  con condición C sin ingresar a vigilancia.
- Botón de descarga a Excel (mismo patrón `io.BytesIO` + `st.download_button`, con `key=` único).

**c) Brecha inversa (un bloque chico, no protagonista)**
CT **con nómina VS pero sin condición registrada**. Son 124 en el universo (100 en el plan): gente en
vigilancia de la salud cuyo CT no tiene Identificación Avanzada registrada en ISTProd. Es un hallazgo
real de calidad de datos, vale la pena mostrarlo aunque sea como métrica + expander con la tabla.

**d) Cobertura por región y por ergónomo**
Barras horizontales o tabla: pendientes por Región y por Ergonomo, para repartir la gestión.
Reutiliza la paleta que ya está en el archivo: `#C0392B` crítica, `#E67E22` no crítica,
`#1A936F` aceptable/ok, `#2E86AB` azul institucional.

**NO repitas** el bloque "🌡️ Vigilancia Ambiental: Identificación y Evaluación" que ya vive dentro
del tab "Estado Seguimiento" (plazos de 90 días de las Críticas y vigencia de 36 meses de las
Aceptables). Ese tab habla del **ambiente**; el nuevo habla de las **personas**. Si algo se solapa,
enlaza conceptualmente con una `st.caption`, no dupliques el cálculo.

---

## 3. CÓMO INTEGRARLO (convenciones del archivo)

- El tab se agrega en `st.tabs([...])` (~línea 977) y en el desempaque
  `tab1, tab2, tab_seg, tab_ind = tabs`. Súmalo al final: `"🩺 Vigilancia de la Salud"`.
- **Usa `df_seg`, NO `df_seg_raw`.** `df_seg` es el que ya trae aplicados los filtros del sidebar
  (ergónomo, gerencia, holding, empleador, región, solo EP, solo CT activos) en el bloque
  "APLICAR FILTROS" (~líneas 905-934). El tab nuevo tiene que respetar esos filtros como todos los demás.
- Guarda el patrón de defensa que usa todo el archivo: `if df_seg.empty: st.info(...)` y
  `if 'columna' in df_seg.columns` antes de tocar cualquier columna. Si falta una columna, el tab
  debe degradar con un mensaje, no reventar.
- **Fechas siempre `dd-mm-aaaa`** al mostrar (`.dt.strftime('%d-%m-%Y').fillna('')`), regla de la casa.
- Reutiliza los helpers que ya existen: `es_ct_activo()`, `parsear_fecha_flexible()`. No los redefinas.
- Prefijo `_` para las variables locales del tab, como en el resto del archivo, para no chocar con el
  espacio de nombres global del script.

---

## 4. CRITERIO DE ACEPTACIÓN (cifras verificadas al 26-08-2026, sin filtros del sidebar)

Con la data del plan (5.500 CT programados; universo completo 7.239) y la ventana VS 2025-2026:

| Métrica | Plan | Universo |
|---|---:|---:|
| CT con condición C o M | **62** (37 C + 25 M) | 72 (45 C + 27 M) |
| … de ellos SIN nómina VS | **22** | 27 |
| … con nómina VS | **40** | 45 |
| …… completos (eval ≥ deben) | 9 | 10 |
| …… incompletos (eval < deben) | 31 | 35 |
| Trabajadores que deben ingresar / evaluados (en CT C-M) | 1.734 / 1.086 | 1.767 / 1.088 |
| CT con nómina VS (total) | 143 | 172 |
| … de esos, sin condición registrada (brecha inversa) | 100 | 124 |

Si tus números no dan esto, el filtro o el manejo de nulos está mal. **Párate y avísame.**

Dato conocido, no es bug: 6 CT (`76726790-811`, `77063395-81`, `78637890-71`, `78973230-29`,
`84865000-5294`, `96856610-515`) tienen vigilancia de la salud pero no están en el universo (no están
en el plan ni tienen actividad en SIGECO/ISTProd), así que nunca aparecerán en el dashboard. Por eso
178 CT en la fuente → 172 en la planilla.

---

## 5. CÓMO QUIERO QUE TRABAJES

- **Antes de inventar un criterio, PREGÚNTAME.** Prefiero una pregunta a un supuesto silencioso.
  Dos que ya sé que vas a necesitar:
  1. **¿Cuál es el plazo normativo para que un CT con condición C o M ingrese a vigilancia de la
     salud?** El bloque de vigilancia ambiental usa 90 días para la *intervención* de las Críticas,
     pero para el ingreso a VS no hay plazo codificado en ninguna parte. Mientras no me preguntes,
     muestra la antigüedad en días como dato, sin semáforo de "vencido".
  2. **¿"Pendiente" en el título del tab incluye los incompletos, o sólo los que no han ingresado?**
     Yo los quiero separados en las métricas; la duda es qué cuenta como el número grande.
- No toques el resto del dashboard: sólo agregas el tab y, si hace falta, helpers nuevos claramente
  delimitados.
- Al terminar, muéstrame las métricas que calculó el tab contra la tabla de aceptación de arriba, y
  un screenshot o la descripción de lo que quedó en pantalla. Si algo no cuadra, dilo.
