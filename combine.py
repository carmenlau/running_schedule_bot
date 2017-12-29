import os
from os import path
import json
os.chdir("output")
files = [f for f in os.listdir(".") if path.isfile(f) and f.endswith('.json') and '_' in f]
d = []
for f in files:
    j = json.loads(open(f, 'r').read())
    d += j

with open('all.json', 'w') as f:
    f.write(json.dumps(d, indent=4))

