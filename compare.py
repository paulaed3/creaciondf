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

def comparar_celdas(df_ref: pd.DataFrame, df_new: pd.DataFrame):
	# Asegurar mismo orden de columnas
	df_new = df_new[df_ref.columns]
	# Igualar tipos básicos a string donde mezclado para evitar falsos positivos
	# (excepto números que comparamos directos)
	dif_mask = (df_ref.ne(df_new)) & ~(df_ref.isna() & df_new.isna())
	coords = dif_mask.stack()
	if not coords.any():
		return []
	difs = []
	for (idx, col), is_diff in coords.items():
		if is_diff:
			difs.append({
				'row_index': idx,
				'column': col,
				'expected': df_ref.at[idx, col],
				'actual': df_new.at[idx, col]
			})
	return difs

def main():
	parser = argparse.ArgumentParser(description="Compara output.xlsx generado con backup_data/output.xlsx referencia")
	parser.add_argument('--ref', default='backup_data/output.xlsx', help='Archivo de referencia (expected)')
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

	# Alinear índices si 'ID' existe (para comparación semántica)
	if 'ID' in df_ref.columns and 'ID' in df_new.columns:
		# Verificar duplicados
		if df_ref['ID'].is_unique and df_new['ID'].is_unique:
			df_ref = df_ref.set_index('ID').sort_index()
			df_new = df_new.set_index('ID').sort_index()
			print("Comparación alineada por ID")
		else:
			print("Advertencia: IDs duplicados; se compara por posición original")

	difs = comparar_celdas(df_ref, df_new)
	total_celdas = df_ref.shape[0] * df_ref.shape[1]
	if not difs:
		print("Celdas: todas coinciden (sin diferencias)")
	else:
		print(f"Celdas diferentes: {len(difs)} de {total_celdas} ({len(difs)/total_celdas:.4%})")
		for d in difs[:args.max_print]:
			print(f"Fila/Index={d['row_index']} Col='{d['column']}' esperado='{d['expected']}' obtenido='{d['actual']}'")
		if len(difs) > args.max_print:
			print(f"... ({len(difs)-args.max_print} diferencias más no impresas)")
		# Exportar
		pd.DataFrame(difs).to_excel(args.export_diff, index=False)
		print(f"Detalle exportado a {args.export_diff}")

if __name__ == '__main__':
	main()

