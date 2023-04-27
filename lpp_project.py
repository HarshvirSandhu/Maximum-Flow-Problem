import openpyxl
import networkx as nx
import numpy as np
import matplotlib.pyplot as plt
from pulp import *

w = openpyxl.open('adj_matrix.xlsx')
mat = []
for idx, cell in enumerate(w['Sheet3']):
    row = []
    for val in cell:
        row.append(val.value)
    if set(row) == {0}:
        n = idx
    mat.append(row)
n = len(mat) - n
mat = np.array(mat)
directed = nx.DiGraph()
g = nx.from_numpy_array(mat, create_using=directed)
pos = nx.spring_layout(g)
nx.draw_networkx(g, pos=pos)
labels = nx.get_edge_attributes(g, 'weight')
nx.draw_networkx_edge_labels(g, pos=pos, edge_labels=labels)
plt.show()


path_dict = {}
for k, v in labels.items():
    path_dict['X'+str(k[0])+'_'+str(k[-1])] = v

prob = LpProblem('Max_Flow', pulp.const.LpMaximize)

path_vars = LpVariable.dicts('Paths', list(path_dict.keys()), 0)
# print(path_vars)
# print(path_dict)
print('Number of variables: ', len(path_dict))
obj_vars = [k for k in path_vars.keys() if k[1] == '0']

prob += (lpSum([path_vars[i] for i in path_vars if i in obj_vars]), 'Objective value')

for path in path_dict:
    prob += (
        lpSum([path_vars[path]]) <= path_dict[path],
        f'Path_limit_{path}'
    )

# print('-----------------------------------------------------')
# print([i.split('_')[-1] for i in path_vars])
# print(len(mat))
# print([path_vars[i] for i in path_vars if int(i.split('_')[0].strip('X')) == int(len(mat)-1)])
# print([int(i.split('_')[0].strip('X')) for i in path_vars])
# print('-----------------------------------------------------')

for node in list(g.nodes)[1:-(n+1)]:
    prob += (
        lpSum([path_vars[i] for i in path_vars if int(i.split('_')[-1]) == int(node)]) == lpSum([path_vars[i] for i in path_vars if int(i.split('_')[0].strip('X')) == int(node)]),
        f'Flow conservation_{node}'
    )

prob += (
    lpSum([path_vars[i] for i in path_vars if int(i.split('_')[-1]) == int(len(mat)-1)]) == lpSum([path_vars[i] for i in path_vars if int(i.split('_')[0].strip('X')) == int(0)]),
    'final_conservation'
)


print(prob)
prob.solve()

print("Status: ", LpStatus[prob.status])
for v in prob.variables():
    print(v.name, "=", v.varValue)

print("Max = ", value(prob.objective))
