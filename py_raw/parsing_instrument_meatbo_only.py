# Парсит из товаров на сайте интсрумента только метоабовские позиции
# выводит экселевский файл со ссылкой, артикулом и названием

import pandas as pd


instr_df = pd.read_excel('./../urls/instrument_all_items.xlsx', engine='xlrd')
# output_df = instr_df.query('название == "Аккумуляторный фонарик Bosch Pro GLI 12V-300"')
# output_df = instr_df[instr_df.название == 'Аккумуляторный фонарик Bosch Pro GLI 12V-300']
output_df = instr_df[instr_df['название'].str.contains("Metabo")]
print(output_df)
writer = pd.ExcelWriter('./../xlsx/instrument_metabo_items_only.xlsx', engine='xlsxwriter')
output_df.to_excel(writer, sheet_name='Sheet1', index=True)
writer.save()
