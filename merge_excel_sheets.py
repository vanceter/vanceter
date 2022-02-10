# 2022/02/10 Terry Vance
# importing the module
import pandas

# NOTE each of the underlying xls files needs to have the PSLC value - in some of them, the column header needs to be renamed from PS Loc
# Also need to make sure you export OpsTracker files with the file name option checked
 
# reading the entire files - commented out to use the commands below instead, pulling only the columns needed
# f_sites = pandas.read_excel("opstracker_sites.xlsx")
# f_gens = pandas.read_excel("opstracker_generators.xlsx")
# f_cells = pandas.read_excel("NorCal_CellInfo.xlsx")
# f_cells5g = pandas.read_excel("norcal_cell_info_5g.xlsx")
# f_pge = pandas.read_excel("PSPS_FIRE_TIER.xlsx")

# reading only the columns needed from each file
# documentation on pandas read_excel https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
f_sites = pandas.read_excel("opstracker_sites.xlsx",usecols=['SITE_NAME','ADDRESS','CITY','COUNTY','PSLC','POWER_METER', 'GEN_STATUS','GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'IS_HUB','IS_HUB_MICROWAVE','REMOTE_MONITORING','SITETECH_NAME','SITETECH_MANAGER_NAME'])
f_gens = pandas.read_excel("opstracker_generators.xlsx", usecols=['PSLC', 'FUEL_TYPE1'])
f_cells = pandas.read_excel("NorCal_CellInfo.xlsx", usecols=['PSLC', 'eNodeB'])
f_cells5g = pandas.read_excel("norcal_cell_info_5g.xlsx", usecols=['PSLC', 'GNODEB'])
f_pge = pandas.read_excel("PSPS_FIRE_TIER.xlsx", usecols=['PSLC', 'Fire Tier', 'PSPS PROB', 'PG&E Fee Property'])
  
# merging the files using PSLC as the index. There are some duplicates in gen and sites files, lots of duplicates in the cell info because of B2B and 5G gNodeBs
# documentation on pandas merge https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.merge.html?highlight=merge#pandas.DataFrame.merge
f_merged = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_pge, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells, left_on="PSLC", right_on="PSLC", how="left")
f_merged = f_merged.merge(f_cells5g, left_on="PSLC", right_on="PSLC", how="left")

# same function but only combining the sites, gens and PSPS/PGE files for Ops
f_merged_ops = f_sites.merge(f_gens, left_on="PSLC", right_on="PSLC", how="left")
f_merged_ops = f_merged_ops.merge(f_pge, left_on="PSLC", right_on="PSLC", how="left")
  
# creating 2 new files, the PSPS_Main for Gennie, and a version of it with eNB/gNB for SP
# documentation on pandas to_excel https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html?highlight=to_excel
f_merged.to_excel("PSPS_MAIN_SP.xlsx", index = False, sheet_name='PSPS_MAIN_SP', columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME', 'eNodeB', 'GNODEB'])
f_merged_ops.to_excel("PSPS_MAIN.xlsx", index = False, sheet_name='PSPS_MAIN', columns=['POWER_METER','Fire Tier', 'PSPS PROB','PSLC', 'PG&E Fee Property', 'SITE_NAME', 'ADDRESS','CITY','COUNTY', 'GEN_STATUS','FUEL_TYPE1', 'GEN_PORTABLE_PLUG', 'GEN_PORTABLE_PLUG_TYPE', 'REMOTE_MONITORING', 'IS_HUB_MICROWAVE', 'IS_HUB','SITETECH_NAME','SITETECH_MANAGER_NAME'])
