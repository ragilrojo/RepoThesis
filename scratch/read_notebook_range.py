import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('strategy_comparison_coba_toGrid2stage.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for i in range(6, 16):
    if i >= len(nb['cells']): break
    cell = nb['cells'][i]
    src = ''.join(cell['source'])
    print(f'=== Cell {i} ({cell["cell_type"]}) ===')
    print(src)
    print()
