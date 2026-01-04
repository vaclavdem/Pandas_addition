import pandas as pd
from services.calculations import build_table_rec

data = pd.read_excel("input_files/task_2_data_ex.xlsx")
results_by_year = {}
columns = ["plant", "fin_mat_id", "fin_mat_rel_type", "fin_mat_prod_type", "fin_prod_quant", "prod_mat_id", \
    "prod_mat_rel_type", "prod_mat_prod_type", "prod_mat_quant", "comp_mat_id", "comp_mat_rel_type", \
    "comp_mat_prod_type", "comp_mat_quant", "year"]

for i in range(len(data["year"])):
    if data["produced_material_release_type"][i] == "FIN":
        plant = data["plant_id"][i]
        fin_mat_id = data["produced_material"][i]
        fin_mat_rel_type = data["produced_material_release_type"][i]
        fin_mat_prod_type = data["produced_material_production_type"][i]
        fin_prod_quant = data["produced_material_quantity"][i]
        year = data["year"][i]
        build_table_rec(data, i, plant, fin_mat_id, fin_mat_rel_type, fin_mat_prod_type, fin_prod_quant,\
                    year, data["produced_material"][i], results_by_year)

with pd.ExcelWriter("output_files/result.xlsx", engine="xlsxwriter") as writer:
    for year, rows in results_by_year.items():
        df = pd.DataFrame(rows, columns=columns)
        df.to_excel(writer, sheet_name=str(year), index=False)