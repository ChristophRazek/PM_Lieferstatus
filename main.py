import pandas as pd
import datetime
import pyodbc
import warnings
from tkinter import messagebox

warnings.simplefilter("ignore")

conn_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(conn_string)

d = datetime.datetime.now()
sql = """SELECT      CASE 
                WHEN STATUS = 1 and (DATEDIFF(day, GETDATE(), [FOB CONF/ LIEFERDATUM]) <= 7) AND ([PE14_MassProdRel] = 0) then 'MPS Fehlen -> Q kontaktieren'
                 WHEN [STATUS] = 1 AND [FOB CONF/ LIEFERDATUM] < SYSDATETIME() THEN 'Veraltetes PM Lieferdatum -> Lieferant kontaktieren' 
				 WHEN STATUS = 1 and [FOB/ WUNSCHDATUM] = [FOB CONF/ LIEFERDATUM] then 'Nicht gepflegt(?) -> PM Lieferanten/Supply Kontaktieren'
				 WHEN STATUS = 3 and ETD is null then 'Fehlende Schmid Daten -> Carmen kontaktieren'
				 WHEN ATD is null and ETD < SYSDATETIME() then 'Fehlendes ATD -> Carmen kontaktieren' 
				 WHEN ATA is null and ETA < SYSDATETIME() then 'Fehlendes ATA -> Carmen kontaktieren' 
				 WHEN [STATUS] in (3,4) AND [LIEFERAVIS SCHMID] < SYSDATETIME() THEN 'Veraltetes Schmid Lieferdatum -> Carmen kontaktieren' 
           
			
			END AS KONTROLLE, DATEDIFF(day, GETDATE(), 
			[FOB CONF/ LIEFERDATUM]) as 'Tage bis Lieferung',
			STATUS,FIXPOSNR, BELEGNR, WARENEINGANGSNR, [SHIPMENT STATUS], ARTIKELNR, BEZEICHNUNG, LIEFERANT, ABFÜLLER, 
            [PREADVISE MENGE], [TRANSPORT MODE],[FOB CONF/ LIEFERDATUM],[PE14_MassProdRel], ETD, ATD, ETA, ATA, [LIEFERAVIS SCHMID], [EXP DISPATCH], [LAST PREADVISE], [COM. EMEA]
			,[FOB/ WUNSCHDATUM]
FROM   db_dataviewer.PE14_SHIPPMENTLIST
WHERE ([SHIPMENT STATUS] <> 'received') AND ([COM. EMEA] IS NULL OR [COM. EMEA] <> 'eRledigt')
order by [FOB CONF/ LIEFERDATUM]"""

df = pd.read_sql_query(sql,conx)
df['WARENEINGANGSNR'] = df['WARENEINGANGSNR'].fillna(0)
df[['BELEGNR', 'WARENEINGANGSNR']] = df[['BELEGNR', 'WARENEINGANGSNR']].astype('int64')
df.dropna(subset=['KONTROLLE'], inplace=True)

df.to_excel(r'S:\EMEA\Kontrollabfragen\PM_Lieferstatus.xlsx', index=False)


with open(r'S:\EMEA\Kontrollabfragen\PM_Lieferstatus.txt', 'w') as f:
    f.write(f'PM Lieferstatus last checked at: {d}')
    f.close()

messagebox.showinfo('Update Erfolgreich!', f'Das PM Status Update wurde am {d} erfolgreich durchgeführt.')