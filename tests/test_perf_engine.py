# -*- coding: utf-8 -*-
import unittest

from perf_engine import calc_perf, find_prefix


class TestPerfEngine(unittest.TestCase):
    def test_find_prefix(self):
        rules = {'2626': {'x': 1}, '262': {'y': 1}}
        self.assertEqual(find_prefix('26260401-03', rules), '2626')

    def test_shortage_pool_equals_gap_times_benchmark(self):
        """缺员总池 = (D-nn)*T/D，每人 (D-nn)*T/(D*nn)。"""
        rules = {
            '2626': {
                '配料基础分': '100',
                '配料定员（人）': '5',
                '配料大清分': '0',
                '过筛分': '0',
            }
        }
        fd = {
            '工序': '配料',
            '日期': '2026-05-04',
            '工时': 8,
            'batches': [{'batch_no': '26260401', 'product_code': '2626', 'box': 0, 'ban': 0}],
            'persons': [
                {'name': 'A', 'coeff': 1.0, 'quality': 0},
                {'name': 'B', 'coeff': 1.0, 'quality': 0},
                {'name': 'C', 'coeff': 1.0, 'quality': 0},
            ],
            'hasDQ': False,
            'hasGS': False,
            'hasZJ': False,
            'qyTotal': 0,
            'xqyTotal': 0,
            'qyPersons': [],
            'zpBatches': 0,
            'zpPersons': [],
            'mgChangePersons': [],
            'mgConfirmPersons': [],
            'jyBaskets': 0,
            'jyPersons': [],
            'fsScore': 0,
            'chong': '',
        }
        out = calc_perf(fd, rules, hour_score=20.0)
        self.assertEqual(len(out), 3)
        T, D, nn = 100.0, 5, 3
        jizhun = T / D
        pool = (D - nn) * jizhun
        share = pool / nn
        for row in out:
            self.assertAlmostEqual(row['系数分'], jizhun * 1.0, places=5)
            self.assertAlmostEqual(row['自动缺员补'], round(share, 2), places=2)
            self.assertIn('缺员2人', row['显示备注'])

    def test_quality_negative_reduces_nc(self):
        rules = {
            '99': {
                '配料基础分': '50',
                '配料定员（人）': '2',
                '配料大清分': '0',
                '过筛分': '0',
            }
        }
        fd = {
            '工序': '配料',
            '日期': '2026-05-04',
            '工时': 8,
            'batches': [{'batch_no': '9901', 'product_code': '99', 'box': 0, 'ban': 0}],
            'persons': [{'name': 'A', 'coeff': 1.0, 'quality': -2}],
            'hasDQ': False,
            'hasGS': False,
            'hasZJ': False,
            'qyTotal': 0,
            'xqyTotal': 0,
            'qyPersons': [],
            'zpBatches': 0,
            'zpPersons': [],
            'mgChangePersons': [],
            'mgConfirmPersons': [],
            'jyBaskets': 0,
            'jyPersons': [],
            'fsScore': 0,
            'chong': '',
        }
        r = calc_perf(fd, rules)[0]
        self.assertLess(r['加减分'], r['系数分'])


if __name__ == '__main__':
    unittest.main()
