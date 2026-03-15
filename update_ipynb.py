import json

with open('strategy_comparison.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for cell in nb.get('cells', []):
    if cell['cell_type'] == 'code':
        new_source = []
        for line in cell['source']:
            # For each line, check if it's plt.show()
            if 'plt.show()' in line:
                cell_text = ''.join(cell['source'])
                
                # Figure 3
                if 'Figure 3 | MST September 2017' in cell_text and 'mst_1.png' not in cell_text:
                    new_source.append("    plt.savefig('mst_1.png', bbox_inches='tight', dpi=300)\n")
                
                # Figure 4
                if 'Figure 4 | MST June 2019' in cell_text and 'mst_2.png' not in cell_text:
                    new_source.append("    plt.savefig('mst_2.png', bbox_inches='tight', dpi=300)\n")
                
                # Figure 8
                if 'FIGURE 8 | Optimal Gamma Selection over Time' in cell_text and 'gamma_plot.png' not in cell_text:
                    if 'FIGURE 9 |' not in line: # guard to prevent multiple insertions in the loop
                        new_source.append("        plt.savefig('gamma_plot.png', bbox_inches='tight', dpi=300)\n")
                
                # Figure 6
                if 'FIGURE 6 | Performances of different portfolio strategies' in cell_text and 'performance_plot.png' not in cell_text:
                    new_source.append("plt.savefig('performance_plot.png', bbox_inches='tight', dpi=300)\n")
                    
            new_source.append(line)
        cell['source'] = new_source

with open('strategy_comparison.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, indent=1)

print("Berhasil menambahkan skrip export gambar ke Jupyter Notebook.")
