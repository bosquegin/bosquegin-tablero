import sys
data = sys.stdin.read()
path = r'C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo\Data\Inventario\gc_stock_stage.txt'
mode = sys.argv[1] if len(sys.argv) > 1 else 'a'
with open(path, mode, encoding='ascii') as f:
    f.write(data)
print(f'wrote {len(data)} chars (mode={mode})', flush=True)
