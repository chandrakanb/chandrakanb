flag = int(input("|\tDryRun2.x:\n|\t\t1.DryRun2.1\n|\t\t2.DryRun2.2\n|\t\t3.DryRun2.3\n|\t\t4.DryRun2.4\n|\tSelect DryRun2.x: "))
cycle = 'DryRun2.1' if flag == 1 else 'DryRun2.2' if flag == 2 else 'DryRun2.3' if flag == 3 else 'DryRun2.4' if flag == 4 else None

print(cycle)
