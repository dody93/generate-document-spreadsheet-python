import gspread
import pandas as pd
from sqlalchemy import create_engine as ce
from docxtpl import DocxTemplate
from pathlib import Path

engine = ce('mysql+pymysql://root:@localhost/db_mahasiswa')
base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
output_dir = base_dir / "OUTPUT"



sh = gspread.service_account()
sa = sh.open('test')

wks = sa.worksheet('Sheet1')
record_data = wks.get_all_records()
records_df = pd.DataFrame.from_dict(record_data)
# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)

sql = '''SELECT * FROM user;'''

df = pd.read_sql(sql,engine)
res = pd.merge(df,records_df)
print(res)
for record in res.to_dict(orient="records"):
    doc = DocxTemplate("my_word.docx")
    doc.render(record)
    output_path = output_dir / f"{record['Nama']}-{record['kelas']}-generated_doc.docx"
    doc.save(output_path)




    




