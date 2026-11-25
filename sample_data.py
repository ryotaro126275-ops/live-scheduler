"""Generate fictional demo data; no real member information."""
from io import BytesIO
from openpyxl import Workbook

def sample_bytes():
    wb=Workbook(); ws=wb.active; ws.title='サンプル'
    ws.append(['メンバー名','学年','','','','','曲数','バンド名','','メンバー1','メンバー2','メンバー3','メンバー4','メンバー5','メンバー6','メンバー7','','出演可能日'])
    for i in range(24):
        row=[None]*18; row[0]=f'架空メンバー{i+1:02}'; row[1]=i%4+1
        if i<12:
            row[6]=i%3+1; row[7]=f'サンプルバンド{i+1:02}'; row[9]=f'架空メンバー{2*i+1:02}'; row[10]=f'架空メンバー{2*i+2:02}'; row[17]=['8,12','13,14,19','8,12,13,14,19'][i%3]
        ws.append(row)
    out=BytesIO(); wb.save(out); return out.getvalue()

if __name__=='__main__':
    from pathlib import Path
    Path('sample_data.xlsx').write_bytes(sample_bytes())
