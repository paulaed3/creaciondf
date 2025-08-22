import pandas as pd
from pathlib import Path
import argparse


def load_excel(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")
    df = pd.read_excel(path)
    # Forzar ID a string si existe
    if 'ID' in df.columns:
        df['ID'] = df['ID'].astype(str)
    return df


def diff_by_id(ref: pd.DataFrame, new: pd.DataFrame):
    if 'ID' not in ref.columns or 'ID' not in new.columns:
        # fallback posicional
        min_rows = min(len(ref), len(new))
        ids = list(range(min_rows))
        ref_use = ref.iloc[:min_rows].copy()
        new_use = new.iloc[:min_rows].copy()
        ref_use['__ID__'] = ids
        new_use['__ID__'] = ids
        id_col = '__ID__'
    else:
        id_col = 'ID'
        ref_use = ref.copy()
        new_use = new.copy()

    # Columnas comunes
    common_cols = [c for c in ref_use.columns if c in new_use.columns]
    # Excluir ID col auxiliar
    if id_col in common_cols:
        pass
    else:
        common_cols.append(id_col)

    # Alinear por ID
    if id_col == 'ID':
        # Ver duplicados
        if ref_use[id_col].duplicated().any() or new_use[id_col].duplicated().any():
            # Expand by cumcount to create unique composite index
            ref_use['__dup__'] = ref_use.groupby(id_col).cumcount()
            new_use['__dup__'] = new_use.groupby(id_col).cumcount()
            ref_use.set_index([id_col, '__dup__'], inplace=True)
            new_use.set_index([id_col, '__dup__'], inplace=True)
        else:
            ref_use.set_index(id_col, inplace=True)
            new_use.set_index(id_col, inplace=True)
    else:
        ref_use.set_index(id_col, inplace=True)
        new_use.set_index(id_col, inplace=True)

    # Unir índices
    all_index = ref_use.index.union(new_use.index)
    ref_al = ref_use.reindex(all_index)
    new_al = new_use.reindex(all_index)

    diffs = []

    # IDs faltantes / sobrantes
    missing_in_new = ref_al.index.difference(new_al.dropna(how='all').index)
    missing_in_ref = new_al.index.difference(ref_al.dropna(how='all').index)

    # Comparar celdas
    for col in common_cols:
        if col == id_col or col == '__dup__':
            continue
        if col not in ref_al.columns or col not in new_al.columns:
            continue
        ref_col = ref_al[col]
        new_col = new_al[col]
        # Diferencias lógicas
        mask = (ref_col != new_col) & ~(ref_col.isna() & new_col.isna())
        if not mask.any():
            continue
        for idx in ref_al.index[mask]:
            id_value = idx if isinstance(idx, tuple) else (idx,)
            # Primer elemento del índice es el ID real
            real_id = id_value[0]
            diffs.append({
                'ID': real_id,
                'COLUMN': col,
                'EXPECTED': ref_col.loc[idx],
                'ACTUAL': new_col.loc[idx]
            })

    # Añadir registros de filas enteras ausentes
    for idx in missing_in_new:
        real_id = idx[0] if isinstance(idx, tuple) else idx
        diffs.append({'ID': real_id, 'COLUMN': '__ROW__', 'EXPECTED': 'ROW_PRESENT_IN_REF', 'ACTUAL': 'MISSING_IN_NEW'})
    for idx in missing_in_ref:
        real_id = idx[0] if isinstance(idx, tuple) else idx
        diffs.append({'ID': real_id, 'COLUMN': '__ROW__', 'EXPECTED': 'MISSING_IN_REF', 'ACTUAL': 'ROW_PRESENT_IN_NEW'})

    return pd.DataFrame(diffs)


def main():
    parser = argparse.ArgumentParser(description='Genera CSV con diferencias por ID y columna entre referencia y nuevo.')
    parser.add_argument('--ref', default='backup_data/output_expect.xlsx', help='Archivo referencia (expected)')
    parser.add_argument('--new', default='output.xlsx', help='Archivo generado (actual)')
    parser.add_argument('--out', default='diff_ids.csv', help='Archivo CSV de salida')
    parser.add_argument('--limit', type=int, default=0, help='Si >0, limitar filas exportadas')
    args = parser.parse_args()

    ref = load_excel(Path(args.ref))
    new = load_excel(Path(args.new))

    df_diffs = diff_by_id(ref, new)
    if args.limit > 0:
        df_diffs = df_diffs.head(args.limit)

    df_diffs.to_csv(args.out, index=False)
    print(f"Diferencias: {len(df_diffs)} filas exportadas a {args.out}")
    # Resumen rápido
    if not df_diffs.empty:
        top_cols = df_diffs['COLUMN'].value_counts().head(10)
        print('Top columnas con diferencias:')
        print(top_cols.to_string())
    else:
        print('No se encontraron diferencias.')


if __name__ == '__main__':
    main()
