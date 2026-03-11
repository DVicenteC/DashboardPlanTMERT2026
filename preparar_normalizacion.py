"""
preparar_normalizacion.py
=========================
Script de dos fases para normalizar ocupaciones y tareas del archivo TMERT.
Todo el flujo usa archivos Excel, sin archivos CSV intermedios.

Algoritmo de clustering: process.dedupe() de thefuzz
  · Agrupa variantes similares en clusters
  · Elige UN representante canónico por cluster (sin mapeos circulares)
  · Solo propone cambios donde variante != canónico

FASE 1 — Generar propuestas (siempre se ejecuta):
  Exporta un único Excel de revisión:
    normalizacion/propuestas_normalizacion.xlsx
    · Hoja "ocupaciones" : variante | canonical_propuesto | similitud
    · Hoja "tareas"      : ídem
  Solo aparecen filas donde se propone un cambio.
  Los valores únicos (sin similar) no aparecen — ya son canónicos.

FASE 2 — Aplicar y exportar Excel corregido:
  Lee las hojas del Excel de propuestas y genera:
    normalizacion/TMERT_2026_normalizado.xlsx
  con las columnas 'ocupaciones' y 'tareas' normalizadas.
  Este archivo es el que se sube a Google Sheets para el dashboard.

Uso:
  python preparar_normalizacion.py

  Primera ejecución → genera propuestas_normalizacion.xlsx
  Tras revisar      → vuelve a ejecutar → genera TMERT_2026_normalizado.xlsx

Requiere: pandas, openpyxl, thefuzz, python-Levenshtein
"""

import re
import unicodedata
import os
import pandas as pd
from thefuzz import fuzz, process

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
EXCEL_PATH = (
    r"C:\EspecialidadesTecnicas\02_PlanAnual_SUSESO\2026"
    r"\Anexo 7.1 Planificaci" + "ó" + r"n Programa de Trabajo TMERT 2026 - Detalles_reformulado.xlsx"
)
SKIPROWS         = 3
UMBRAL_SIMILITUD = 88   # % mínimo para agrupar variantes (≥88 reduce falsos positivos)
OUTPUT_DIR       = os.path.join(os.path.dirname(__file__), "normalizacion")

# Archivos de trabajo (todo Excel)
EXCEL_PROPUESTAS = os.path.join(OUTPUT_DIR, "propuestas_normalizacion.xlsx")
EXCEL_SALIDA     = os.path.join(OUTPUT_DIR, "TMERT_2026_normalizado.xlsx")


# ── UTILIDADES DE TEXTO ───────────────────────────────────────────────────────
def quitar_tildes(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_texto(texto: str) -> str:
    """Minúsculas, sin tildes, sin doble espacio, sin caracteres raros."""
    t = str(texto).strip().lower()
    t = quitar_tildes(t)
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"[^a-z0-9 /().-]", "", t)
    return t.strip()


def extraer_atomicos(serie: pd.Series,
                     sep_folio: str = "||",
                     sep_interno: str = ",") -> list[str]:
    """Extrae valores únicos atómicos normalizados de una columna multi-valor."""
    valores = set()
    for celda in serie.dropna().astype(str):
        for por_folio in celda.split(sep_folio):
            for item in por_folio.split(sep_interno):
                v = normalizar_texto(item)
                if v:
                    valores.add(v)
    return sorted(valores)


# ── FASE 1: CLUSTERING CON DEDUPE → EXCEL DE PROPUESTAS ───────────────────────
def construir_tabla_con_dedupe(valores: list[str],
                                umbral: int = UMBRAL_SIMILITUD) -> pd.DataFrame:
    """
    Usa process.dedupe() para clusterizar variantes sin mapeos circulares.

    Lógica:
      1. dedupe() agrupa valores similares y elige UN canónico por cluster.
      2. Para cada variante que no es canónica, se busca su canónico con extractOne().
      3. Solo se incluyen en el Excel las filas donde variante != canonical_propuesto.

    Ventaja sobre extract() pairwise: elimina el problema A→B / B→A (circulares).
    """
    print(f"  Clusterizando {len(valores)} valores únicos (umbral={umbral}%)...")

    # Paso 1: obtener un canónico por cluster
    canonicos = sorted(
        process.dedupe(valores, threshold=umbral, scorer=fuzz.token_sort_ratio)
    )
    canonicos_set = set(canonicos)
    n_variantes = len(valores) - len(canonicos)
    print(f"  → {len(canonicos)} clusters  |  {n_variantes} variantes a proponer cambio")

    # Paso 2: mapear variantes → su canónico
    filas = []
    for i, val in enumerate(valores):
        if i % 200 == 0 and i > 0:
            print(f"    {i}/{len(valores)}...")

        if val in canonicos_set:
            # Este valor ya ES el representante de su cluster → no cambia
            continue

        resultado = process.extractOne(val, canonicos, scorer=fuzz.token_sort_ratio)
        if resultado and resultado[1] >= umbral:
            canonical, score = resultado
            filas.append({
                "variante":            val,
                "canonical_propuesto": canonical,
                "similitud":           score,
            })

    df = pd.DataFrame(filas) if filas else pd.DataFrame(
        columns=["variante", "canonical_propuesto", "similitud"]
    )
    if not df.empty:
        df = df.sort_values(["canonical_propuesto", "similitud"],
                            ascending=[True, False])
    return df


def fase1_generar_propuestas(df_raw: pd.DataFrame):
    """Genera el Excel de propuestas para revisión manual."""

    print("\n[Fase 1 · 1/2] Extrayendo y clusterizando ocupaciones...")
    oc_atomicos = extraer_atomicos(df_raw["ocupaciones"],
                                   sep_folio="||", sep_interno="|")
    print(f"  → {len(oc_atomicos)} ocupaciones únicas")
    df_oc = construir_tabla_con_dedupe(oc_atomicos)

    print("\n[Fase 1 · 2/2] Extrayendo y clusterizando tareas...")
    ta_atomicos = extraer_atomicos(df_raw["tareas"],
                                   sep_folio="||", sep_interno=",")
    print(f"  → {len(ta_atomicos)} tareas únicas")
    df_ta = construir_tabla_con_dedupe(ta_atomicos)

    # Exportar ambas hojas en un único Excel
    with pd.ExcelWriter(EXCEL_PROPUESTAS, engine="openpyxl") as writer:
        df_oc.to_excel(writer, index=False, sheet_name="ocupaciones")
        df_ta.to_excel(writer, index=False, sheet_name="tareas")

        for sheet_name, df in [("ocupaciones", df_oc), ("tareas", df_ta)]:
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df.columns, start=1):
                if df.empty:
                    max_len = len(col_name) + 4
                else:
                    max_len = max(df[col_name].astype(str).str.len().max(),
                                  len(col_name)) + 4
                ws.column_dimensions[
                    ws.cell(row=1, column=col_idx).column_letter
                ].width = min(max_len, 60)

    print("\n══════════════════════════════════════════════════════")
    print("FASE 1 COMPLETA")
    print("══════════════════════════════════════════════════════")
    print(f"  Ocupaciones : {len(oc_atomicos):4d} valores únicos  →  {len(df_oc):3d} cambios propuestos")
    print(f"  Tareas      : {len(ta_atomicos):4d} valores únicos  →  {len(df_ta):3d} cambios propuestos")
    print()
    print("PRÓXIMOS PASOS:")
    print(f"  1. Abre en Excel: {EXCEL_PROPUESTAS}")
    print("  2. En cada hoja, revisa la columna 'canonical_propuesto'")
    print("     · Si la propuesta es correcta  → déjala como está")
    print("     · Si el algoritmo se equivocó  → corrige el canonical_propuesto")
    print("     · Si no debe cambiar nada       → borra esa fila")
    print("  3. Guarda el Excel (mismo nombre y ubicación)")
    print("  4. Vuelve a ejecutar este script → generará el Excel corregido")


# ── FASE 2: APLICAR → EXCEL CORREGIDO ─────────────────────────────────────────
def cargar_mapa_desde_excel(hoja: str) -> dict[str, str]:
    """Lee una hoja del Excel de propuestas y retorna {variante → canonical}."""
    df = pd.read_excel(EXCEL_PROPUESTAS, sheet_name=hoja, dtype=str)
    df = df.dropna(subset=["variante", "canonical_propuesto"])
    return dict(zip(
        df["variante"].str.strip(),
        df["canonical_propuesto"].str.strip()
    ))


def normalizar_celda(celda, mapa: dict[str, str],
                     sep_folio: str = "||", sep_interno: str = ",") -> str:
    """
    Reemplaza cada valor atómico de la celda con su canonical.
    Preserva la estructura de separadores original.
    Valores no encontrados en el mapa se conservan normalizados en MAYÚSCULAS.
    """
    if pd.isna(celda) or str(celda).strip() == "":
        return celda

    folios_resultado = []
    for folio in str(celda).split(sep_folio):
        items_resultado = []
        for item in folio.split(sep_interno):
            item_strip = item.strip()
            if not item_strip:
                items_resultado.append("")
                continue
            clave = normalizar_texto(item_strip)
            canonical = mapa.get(clave, clave)
            items_resultado.append(canonical.upper())
        folios_resultado.append(sep_interno.join(items_resultado))

    return sep_folio.join(folios_resultado)


def fase2_aplicar_y_exportar(df_raw: pd.DataFrame):
    """Lee las propuestas aprobadas y exporta el Excel con datos normalizados."""
    print("\n══════════════════════════════════════════════════════")
    print("FASE 2 — APLICANDO NORMALIZACIÓN")
    print("══════════════════════════════════════════════════════")

    mapa_oc = cargar_mapa_desde_excel("ocupaciones")
    mapa_ta = cargar_mapa_desde_excel("tareas")
    print(f"  Mapa ocupaciones : {len(mapa_oc)} entradas")
    print(f"  Mapa tareas      : {len(mapa_ta)} entradas")

    df_out = df_raw.copy()

    print("  Normalizando 'ocupaciones'...")
    df_out["ocupaciones"] = df_out["ocupaciones"].apply(
        lambda c: normalizar_celda(c, mapa_oc, sep_folio="||", sep_interno="||")
    )

    print("  Normalizando 'tareas'...")
    df_out["tareas"] = df_out["tareas"].apply(
        lambda c: normalizar_celda(c, mapa_ta, sep_folio="||", sep_interno=",")
    )

    print(f"  Exportando → {EXCEL_SALIDA}")
    with pd.ExcelWriter(EXCEL_SALIDA, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="TMERT_normalizado")

    cambios_oc = int((df_out["ocupaciones"].fillna("") != df_raw["ocupaciones"].fillna("")).sum())
    cambios_ta = int((df_out["tareas"].fillna("")     != df_raw["tareas"].fillna("")).sum())
    print()
    print("RESULTADO:")
    print(f"  Filas con ocupaciones modificadas : {cambios_oc}")
    print(f"  Filas con tareas modificadas      : {cambios_ta}")
    print(f"  Excel listo para subir a Google Sheets: {EXCEL_SALIDA}")


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"Leyendo Excel fuente: {EXCEL_PATH}")
    df_raw = pd.read_excel(EXCEL_PATH, skiprows=SKIPROWS)
    print(f"  → {df_raw.shape[0]} filas, {df_raw.shape[1]} columnas")

    # Fase 1 siempre regenera las propuestas
    fase1_generar_propuestas(df_raw)

    # Fase 2 solo si el usuario revisó el Excel de propuestas
    if os.path.exists(EXCEL_PROPUESTAS):
        respuesta = input(
            "\n¿Deseas aplicar las propuestas y generar el Excel corregido? [s/N]: "
        ).strip().lower()
        if respuesta == "s":
            fase2_aplicar_y_exportar(df_raw)
        else:
            print("  Fase 2 omitida. Revisa el Excel de propuestas y vuelve a ejecutar.")
    else:
        print(f"\n  Fase 2 pendiente: revisa {EXCEL_PROPUESTAS} y vuelve a ejecutar.")


if __name__ == "__main__":
    main()
