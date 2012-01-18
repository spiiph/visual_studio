import visual_studio

def print_properties(p, c):
    if p.Properties is not None:
        for r in p.Properties:
            f.write("   " * c + "|* %-40s" % r.Name)
            try:
                f.write("%s\n" % r.Value)
            except:
                f.write("<Exception>\n")

def print_item(p, c):
    print_properties(p, c)
    if p.ProjectItems is not None:
        for item in p.ProjectItems:
            f.write("   " * c + "|- %s\n" % item.Name)
            print_item(item, c+1)

f = open("test.log", "w")

for p in visual_studio.dte.projects:
    f.write("|- %s\n" % p.Name)
    print_item(p, 1)

f.close()
