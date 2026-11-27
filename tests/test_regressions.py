import io
import json
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import fitz
import pandas as pd
from openpyxl import load_workbook
import rules
import engine


def sample(**updates):
    row = {'产品品名':'收纳盒', '产品英文名':'Go Cook Storage Box', 'PO号':'380123',
           'TPNB':'123456789', 'TPND':'012345678', 'ITF':'05063638227053',
           'EAN':'5063638226940', 'category':'Cook & Dining', '总外箱数':'10',
           '外箱装量':'12', '内盒装量':'4', '毛重':'2', '净重':'1',
           '外箱长':'10', '外箱宽':'10', '外箱高':'10'}
    row.update(updates)
    return row


class RuleTests(unittest.TestCase):
    def issues(self, row, selected=None):
        return rules.preflight(pd.DataFrame([row]), selected or rules.ALL_TYPES)

    def test_valid_uk_and_ce(self):
        self.assertEqual(self.issues(sample()), [])
        self.assertEqual(self.issues(sample(PO号='999123', CEORMSNO='2005101001984', SKU='123456789')), [])

    def test_fractional_ratio_is_rejected(self):
        self.assertTrue(self.issues(sample(外箱装量='10'), ['BD']))
        with self.assertRaises(ValueError): rules.derive_bd_barcode('05063638227053', '10', '4')

    def test_itf_quantities(self):
        for qty in ['0', '-1', 'abc1', '1.2', 'NaN', 'Infinity']:
            with self.subTest(qty=qty): self.assertTrue(self.issues(sample(内盒装量=qty), ['ITF']))

    def test_equal_invalid_quantities_do_not_bypass_validation(self):
        for qty in ['0','NaN','abc1','-2']:
            with self.subTest(qty=qty):
                self.assertTrue(self.issues(sample(外箱装量=qty,内盒装量=qty), ['BD']))

    def test_optional_inner_and_equal_quantity(self):
        for qty in ['', '12']:
            row = sample(内盒装量=qty)
            self.assertEqual(self.issues(row), [])
            self.assertEqual(rules.requested_kinds(row, rules.ALL_TYPES), ['outer','itf_uk'])

    def test_required_fields_depend_on_label(self):
        row = {'产品品名':'收纳盒','产品英文名':'Box','PO号':'380','ITF':'05063638227053',
               'TPNB':'123456789','TPND':'012345678','外箱装量':'12'}
        self.assertEqual(self.issues(row, ['ITF']), [])
        self.assertTrue(self.issues(row, ['外箱']))
        self.assertTrue(self.issues({**row,'TPNB':'/'}, ['ITF']))

    def test_stale_formula_values_recomputed(self):
        row = rules.enrich_row(sample(订单国家='CE', 外箱尺寸='old', BD条码号='old',
                                     双人抬标识='双人抬', 外箱文件命名列='old', 有CDU加印文字='old'))
        self.assertEqual(row['订单国家'], 'UK')
        self.assertEqual(row['外箱尺寸'], '10*10*10cm')
        self.assertEqual(row['BD条码号'], '02050636382270533703')
        self.assertEqual(row['双人抬标识'], '')
        self.assertNotIn('外箱文件命名列', row)
        self.assertNotIn('有CDU加印文字', row)

    def test_same_name_different_codes_never_dedup(self):
        rows = pd.DataFrame([sample(),sample(ITF='05063446647500')])
        plan = rules.build_plan(rows, ['内盒','ITF','BD'])
        self.assertEqual(len(plan), 6)
        self.assertEqual(len({p['name'] for p in plan}),6)
        self.assertTrue(all(p['duplicate_of'] is None for p in plan))

    def test_content_changes_same_code_never_dedup(self):
        rows = pd.DataFrame([sample(),sample(内盒装量='6')])
        plan = rules.build_plan(rows, ['BD'])
        self.assertNotEqual(plan[0]['name'],plan[1]['name'])

    def test_identical_records_merge(self):
        plan = rules.build_plan(pd.DataFrame([sample(),sample()]),['BD'])
        self.assertEqual(plan[1]['duplicate_of'],2)
        self.assertEqual(plan[0]['name'],plan[1]['name'])

    def test_fingerprint_changes_for_all_inputs(self):
        df = pd.DataFrame([sample()])
        base = rules.input_fingerprint(df,['ITF'],{})
        self.assertNotEqual(base,rules.input_fingerprint(pd.DataFrame([sample(内盒装量='6')]),['ITF'],{}))
        self.assertNotEqual(base,rules.input_fingerprint(df,['BD'],{}))
        self.assertNotEqual(base,rules.input_fingerprint(df,['ITF'],{'05063638227053':('Y','')}))

    def test_explicit_exclusions(self):
        row = sample(PO号='', 是否需要ITF_UK='')
        row.update({'是否需要ITF-UK':'N','是否需要ITF-CE':'N'})
        self.assertEqual(self.issues(row,['ITF']),[])
        self.assertEqual(rules.build_plan(pd.DataFrame([row]),['ITF']),[])

    def test_cdu_conflicts(self):
        df = pd.DataFrame([['05063638227053','Y',''],['05063638227053','N','']],columns=rules.CDU_COLS)
        with self.assertRaises(ValueError): rules.cdu_df_to_map(df)

    def test_blank_template_has_no_formulas_or_sample_orders(self):
        wb = load_workbook(io.BytesIO(engine.build_blank_template()))
        ws = wb['数据源']
        self.assertTrue(all(c.data_type != 'f' for row in ws for c in row))
        self.assertTrue(all(c.value is None for c in ws[2]))
        self.assertNotIn('BD条码号',[c.value for c in ws[1]])
        self.assertEqual(ws['G501'].number_format,'@')

    def test_invalid_weight_and_switch(self):
        self.assertTrue(self.issues(sample(净重='3')))
        self.assertTrue(self.issues(sample(是否要做BD='maybe')))


class BatchTests(unittest.TestCase):
    @staticmethod
    def write_pdf(row, work, out, profile):
        doc=fitz.open(); page=doc.new_page(width=400,height=150)
        page.insert_text((20,40),'Test label')
        doc.save(Path(out, rules.resolved_name('bd',row)+'.pdf'))
        doc.close()

    def test_validation_runs_before_rendering(self):
        with patch.object(engine,'_install_fonts') as fonts:
            with self.assertRaises(ValueError): engine.run_all(pd.DataFrame([sample(内盒装量='0')]),['BD'])
            fonts.assert_not_called()

    def test_failed_first_duplicate_retries_and_packages_only_success(self):
        calls=[]
        def render(*args):
            calls.append(1)
            self.write_pdf(*args)
            if len(calls)==1: raise RuntimeError('conversion failed after writing')
        with patch.object(engine,'_install_fonts'), patch.object(engine,'gen_bd',side_effect=render):
            batch=engine.run_all(pd.DataFrame([sample(),sample(),sample()]),['BD'])
        self.assertEqual([x['状态'] for x in batch['results']],['失败','成功','跳过'])
        self.assertEqual(len(batch['pdfs']),1)
        with zipfile.ZipFile(io.BytesIO(batch['zip_data'])) as archive:
            self.assertEqual(sum(n.endswith('.pdf') for n in archive.namelist()),1)
            self.assertIn('生成明细.csv',archive.namelist())
            self.assertEqual(json.loads(archive.read('批次信息.json'))['failed'],1)

    def test_entirely_failed_job_excludes_partial_pdf(self):
        def fail(*args):
            self.write_pdf(*args)
            raise RuntimeError('failure')
        with patch.object(engine,'_install_fonts'),patch.object(engine,'gen_bd',side_effect=fail):
            batch=engine.run_all(pd.DataFrame([sample()]),['BD'])
        self.assertEqual(batch['pdfs'],{})


class PageTests(unittest.TestCase):
    def test_page_smoke_and_stale_batch(self):
        from streamlit.testing.v1 import AppTest
        at=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py')).run()
        self.assertFalse(at.exception)
        at.session_state['input_df']=pd.DataFrame([sample(内盒装量='0')])
        at.run()
        self.assertFalse(at.exception)
        generate=next(b for b in at.button if b.label=='🚀 开始生成')
        self.assertTrue(generate.disabled)
        at.session_state['batch']={'fingerprint':'old'}
        at.run()
        self.assertIsNone(at.session_state['batch'])
        self.assertFalse(at.exception)


if __name__=='__main__': unittest.main()
