import sys, pandas
sys.path.insert(1,'/net/users/chenghua/Python/Util')

def readSheet(xl, sheet_name, column):
    df1 = xl.parse(sheet_name)
    try:
        l_ccd_ids = df1[column]
    except KeyError:
        return ([],{},[])
    l_header = []
    for each in df1.columns:
        l_header.append(str(each))
        
    l_ = []
    d_ = {}
    for index, row in df1.iterrows():
        id = row[column]
        ## print id
        if pandas.notna(id):
            ## d_[id] = row # works, but the row object is not standard dictionary.
            id = str(id)
            l_.append(id)

            d_row = {}
            for key in df1.columns:
                ## print key
                if pandas.notna(row[key]):
                    ## print row[key]
                    try:
                        d_row[str(key)] = str(row[key])
                    except UnicodeEncodeError, msg:
                        print key, row[key]
                        print msg
                else:
                    d_row[str(key)] = ""
            d_[id] = d_row
            
    return (l_,d_,l_header)

def main():
    ## filepath = "/net/users/chenghua/Projects/Carb/CCD/SNFG/SNFG_parent_core.xlsx"
    ## xl = pandas.ExcelFile(filepath)
    ## for sheet_name in xl.sheet_names:
    ##     print sheet_name
    ##     if sheet_name != "Misc":
    ##         (l_,d_,l_header) = readSheet(xl,sheet_name, "pdb_ligand")
    ##         for id in l_:
    ##             print id
    
    filepath = "/net/users/chenghua/Projects/Carb/CCD/Update_SNFG_mod/Review_ligands/1st_batch_mod_clean.xlsx"
    xl = pandas.ExcelFile(filepath)
    for sheet_name in xl.sheet_names:
        print sheet_name

    (l_,d_,l_header) = readSheet(xl,"Updated", "pdb_ligand")
    for id in l_:
        print id
    print d_["BOG"]

    
    ## print xl.sheet_names
    ## l_ids = []
    ## for sheet_name in xl.sheet_names:
    ##     print sheet_name
    ##     if sheet_name != "Misc":
    ##         (l_,d_,l_header) = readSheet(xl,sheet_name)
    ##         if l_:
    ##             for id in l_:
    ##                 l_ids.append(id)
    ##                 ## print id, d_[id]["type_carbon"], d_[id]["type_acetyl"], d_[id]["type_ring"], d_[id]["mod"]
    ##                 l_template = [d_[id]["type_carbon"], d_[id]["type_acetyl"].capitalize()[:4]]
    ##                 if d_[id]["type_ring"]=="p":
    ##                     l_template.append("Pyranose")
    ##                 elif d_[id]["type_ring"]=="f":
    ##                     l_template.append("Furanose")
    ##                 else:
    ##                     l_template.append("")
    ##                 if d_[id]["mod"].strip():
    ##                     l_template.append(d_[id]["mod"].strip())
    ##                 filename_template = "-".join(l_template) + ".cif"
    ##                 filepath_template = os.path.join(folder_template, filename_template)
    ##                 if os.path.isfile(filepath_template):
    ##                     pass
    ##                     #print id, filename_template
    ##                 else:
    ##                     print id
    ## print len(l_ids)

if __name__=='__main__':
	main()
