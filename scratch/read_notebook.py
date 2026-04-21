import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('strategy_comparison_coba_toGrid2stage.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for i, cell in enumerate(nb['cells']):
    src = ''.join(cell['source'])
    if src.strip():
        print(f'=== Cell {i} ({cell["cell_type"]}) ===')
        print(src[:3000])
        print()
