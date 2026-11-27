"""Real DOCX → PDF integration test. Run from repository root."""
import json
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import pandas as pd
import fitz
import zxingcpp
from PIL import Image
from engine import run_all, preview_pdf
from rules import ALL_TYPES, build_plan
from test_regressions import sample

rows=[sample(),sample(PO号='999123',CEORMSNO='2005101001984',SKU='123456789'),
      sample(PO号='520123',内盒装量='',产品品名='长品名',产品英文名='F&F Home Acacia Wooden Tray with 3 bowls black'),
      sample(ITF='05063446647500',EAN='5063446647500')]
df = pd.DataFrame(rows)
plan = build_plan(df, ALL_TYPES)
expected = {f"{p['type']}/{p['name']}.pdf": p for p in plan}
batch=run_all(df,ALL_TYPES)
out=Path('output/smoke');out.mkdir(parents=True,exist_ok=True)
(out/'manifest.json').write_text(json.dumps(batch['manifest'],indent=2))
(out/'results.json').write_text(json.dumps(batch['results'],ensure_ascii=False,indent=2))
print(json.dumps(batch['results'],ensure_ascii=False,indent=2))
for i,(name,data) in enumerate(batch['pdfs'].items()):
    (out/f'{i:02d}.pdf').write_bytes(data)
    (out/f'{i:02d}.png').write_bytes(preview_pdf(data))
    with fitz.open(stream=data,filetype='pdf') as pdf:
        assert len(pdf)==1,name
        text = pdf[0].get_text()
        assert '{{' not in text,name
        item = expected[name]
        if item['kind'] in ('itf_uk', 'itf_ce', 'bd'):
            pix = pdf[0].get_pixmap(matrix=fitz.Matrix(3, 3))
            image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
            codes = zxingcpp.read_barcodes(image)
            assert len(codes) == 1, (name, codes)
            row = item['row']
            if item['kind'] == 'bd':
                value = row['BD条码号']
                assert codes[0].symbology_identifier == ']C1', name
                assert codes[0].text == f"(02){value[2:16]}(37){value[18:]}", name
            else:
                assert codes[0].text == row['ITF'], name
assert batch['manifest']['failed']==0,batch['manifest']
assert batch['manifest']['pdf_count']==12,batch['manifest']
(out/'batch.zip').write_bytes(batch['zip_data'])
print('Integration passed:',batch['manifest'])
