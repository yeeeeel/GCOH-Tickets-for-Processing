import pandas as pd
import networkx as satan

print("running")
df = pd.read_excel('pandas.xlsx')


G = satan.from_pandas_edgelist(df, 'Parent', 'Node', create_using=satan.DiGraph)
roots = [v for v, d in G.in_degree() if d == 0]
leaves = [v for v, d in G.out_degree() if d == 0]

all_paths = []
for root in roots:
    for leaf in leaves:
        paths = satan.all_simple_paths(G, root, leaf)
        all_paths.extend(paths)

for node in satan.nodes_with_selfloops(G):
    all_paths.append([node, node])

output = pd.DataFrame(sorted(all_paths)).add_prefix('level').fillna(' ')

output.to_excel('pandas.xlsx',sheet_name='Output')

#use 0a[Blank] this will solve ur problems
print("its joeber")