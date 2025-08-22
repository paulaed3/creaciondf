import pandas as pd
from pathlib import Path
import argparse

def cargar_excel(path: Path):
	if not path.exists():
		raise FileNotFoundError(f"No existe el archivo: {path}")
	return pd.read_excel(path)

def columnas_diferentes(cols_ref, cols_new):
	set_ref = set(cols_ref)
	set_new = set(cols_new)
	faltan = [c for c in cols_ref if c not in set_new]
	sobrantes = [c for c in cols_new if c not in set_ref]
	orden_difiere = cols_ref != cols_new
	return faltan, sobrantes, orden_difiere

def comparar_por_id(df_ref: pd.DataFrame, df_new: pd.DataFrame, id_col: str = 'ID'):
	"""Compara celda a celda alineando por ID (maneja duplicados) o, si no existe, por posición.
	Devuelve lista de dicts: ID, COLUMN, EXPECTED, ACTUAL."""
	if id_col in df_ref.columns and id_col in df_new.columns:
		ref_use = df_ref.copy()
		new_use = df_new.copy()
		# Manejo de duplicados
		if ref_use[id_col].duplicated().any() or new_use[id_col].duplicated().any():
			ref_use['__dup__'] = ref_use.groupby(id_col).cumcount()
			new_use['__dup__'] = new_use.groupby(id_col).cumcount()
			ref_use.set_index([id_col, '__dup__'], inplace=True)
			new_use.set_index([id_col, '__dup__'], inplace=True)
		else:
			ref_use.set_index(id_col, inplace=True)
			new_use.set_index(id_col, inplace=True)
		index_name = id_col
	else:
		# Fallback posicional
		ref_use = df_ref.copy()
		new_use = df_new.copy()
		ref_use['__POS__'] = range(len(ref_use))
		new_use['__POS__'] = range(len(new_use))
		ref_use.set_index('__POS__', inplace=True)
		new_use.set_index('__POS__', inplace=True)
		index_name = '__POS__'

	# Unir índices para detectar filas faltantes/sobrantes
	all_index = ref_use.index.union(new_use.index)
	ref_al = ref_use.reindex(all_index)
	new_al = new_use.reindex(all_index)

	# Columnas comunes (mismo orden de referencia) excluyendo la columna de ID (se usará como índice)
	common_cols = [c for c in df_ref.columns if c in df_new.columns and c != id_col]

	difs = []
	for col in common_cols:
		ref_col = ref_al[col]
		new_col = new_al[col]
		mask = (ref_col != new_col) & ~(ref_col.isna() & new_col.isna())
		if not mask.any():
			continue
		for idx in ref_al.index[mask]:
			# idx puede ser tupla (ID, dup) o simple
			real_id = idx[0] if isinstance(idx, tuple) else idx
			difs.append({
				'ID': real_id,
				'COLUMN': col,
				'EXPECTED': ref_col.loc[idx],
				'ACTUAL': new_col.loc[idx]
			})

	# Filas presentes sólo en uno de los dataframes
	missing_in_new = ref_al.index.difference(new_al.dropna(how='all').index)
	missing_in_ref = new_al.index.difference(ref_al.dropna(how='all').index)
	for idx in missing_in_new:
		real_id = idx[0] if isinstance(idx, tuple) else idx
		difs.append({'ID': real_id, 'COLUMN': '__ROW__', 'EXPECTED': 'ROW_PRESENT_IN_REF', 'ACTUAL': 'MISSING_IN_NEW'})
	for idx in missing_in_ref:
		real_id = idx[0] if isinstance(idx, tuple) else idx
		difs.append({'ID': real_id, 'COLUMN': '__ROW__', 'EXPECTED': 'MISSING_IN_REF', 'ACTUAL': 'ROW_PRESENT_IN_NEW'})

	return difs, index_name

def main():
	parser = argparse.ArgumentParser(description="Compara output.xlsx generado con backup_data/output.xlsx referencia")
	parser.add_argument('--ref', default='backup_data/output_expect.xlsx', help='Archivo de referencia (expected)')
	parser.add_argument('--new', default='output.xlsx', help='Archivo generado (actual)')
	parser.add_argument('--export-diff', default='diff_cells.xlsx', help='Archivo Excel donde exportar celdas diferentes')
	parser.add_argument('--max-print', type=int, default=50, help='Máximo de diferencias a imprimir en consola')
	args = parser.parse_args()

	ref_path = Path(args.ref)
	new_path = Path(args.new)
	print(f"Leyendo referencia: {ref_path}")
	df_ref = cargar_excel(ref_path)
	print(f"Leyendo generado : {new_path}")
	df_new = cargar_excel(new_path)

	# Columnas
	faltan, sobrantes, orden_difiere = columnas_diferentes(df_ref.columns.tolist(), df_new.columns.tolist())
	if faltan or sobrantes or orden_difiere:
		print("=== DIFERENCIAS DE COLUMNAS ===")
		if faltan:
			print(f"Columnas faltantes en nuevo ({len(faltan)}): {faltan}")
		if sobrantes:
			print(f"Columnas sobrantes en nuevo ({len(sobrantes)}): {sobrantes}")
		if not faltan and not sobrantes and orden_difiere:
			print("Mismo conjunto de columnas pero el orden difiere")
	else:
		print("Columnas: idénticas (mismo conjunto y orden)")

	# Filas
	if df_ref.shape[0] != df_new.shape[0]:
		print(f"=== DIFERENCIA EN NUMERO DE FILAS === referencia={df_ref.shape[0]} nuevo={df_new.shape[0]}")
	else:
		print(f"Filas: misma cantidad = {df_ref.shape[0]}")

	# Si columnas no coinciden abortar comparación de celdas
	if faltan or sobrantes:
		print("No se compara celda por celda porque difiere el conjunto de columnas")
		return

	# Comparación por ID / posición
	difs, index_name = comparar_por_id(df_ref, df_new, id_col='ID')
	total_celdas = df_ref.shape[0] * len(df_ref.columns)
	if not difs:
		print("Celdas: todas coinciden (sin diferencias)")
	else:
		print(f"Celdas diferentes: {len(difs)} de {total_celdas} ({len(difs)/total_celdas:.4%})")
		for d in difs[:args.max_print]:
			print(f"{index_name}={d['ID']} Col='{d['COLUMN']}' esperado='{d['EXPECTED']}' obtenido='{d['ACTUAL']}'")
		if len(difs) > args.max_print:
			print(f"... ({len(difs)-args.max_print} diferencias más no impresas)")
		# Exportar
		pd.DataFrame(difs).to_excel(args.export_diff, index=False)
		print(f"Detalle exportado a {args.export_diff}")

if __name__ == '__main__':
	main()

