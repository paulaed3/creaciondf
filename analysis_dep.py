import pandas as pd

ref = pd.read_excel('backup_data/output_expect.xlsx',
                    usecols=lambda c: 'DEPRES' in c or 'ANSIED' in c)
new = pd.read_excel(
    'output.xlsx', usecols=lambda c: 'DEPRES' in c or 'ANSIED' in c)
cols = ref.columns.tolist()
print('Columnas:', cols)
report = []
for c in cols:
    if c not in new:
        continue
    diff = (ref[c] != new[c]) & ~(ref[c].isna() & new[c].isna())
    report.append((c, diff.sum(), ref[c].nunique(
        dropna=True), new[c].nunique(dropna=True)))
print('Resumen difs (col, dif_count, nunique_ref, nunique_new):')
for r in report:
    print(r)
# Extras en AMIGO
if 'DEPRESION AMIGO' in ref and 'DEPRESION AMIGO' in new:
    mask_extra = ref['DEPRESION AMIGO'].isna() & new['DEPRESION AMIGO'].notna()
    print('DEPRESION AMIGO valores extras:', mask_extra.sum())
    print(new.loc[mask_extra, 'DEPRESION AMIGO'].value_counts().head())
# Swaps USTED/FAMILIAR
if {'DEPRESION USTED', 'DEPRESION FAMILIAR'} <= set(ref.columns):
    swap_mask = (ref['DEPRESION USTED'] != new['DEPRESION USTED']) & (
        ref['DEPRESION FAMILIAR'] != new['DEPRESION FAMILIAR'])
    print('Posibles swaps depresion USTED/FAMILIAR:', swap_mask.sum())
    print('Ejemplo ref:', ref.loc[swap_mask, [
          'DEPRESION USTED', 'DEPRESION FAMILIAR']].head().to_dict('records'))
    print('Ejemplo new:', new.loc[swap_mask, [
          'DEPRESION USTED', 'DEPRESION FAMILIAR']].head().to_dict('records'))
