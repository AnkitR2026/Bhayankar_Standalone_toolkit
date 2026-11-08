"""Tests for the workflow runner (bhayankar_run.py).

Uses the synthetic fixtures only. The runner is exercised with build=False here; the full build (offline scaffold +
offline validator, no Node.js) is covered in test_bhayankar.py and test_standalone.py.
"""
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.dirname(HERE)
sys.path.insert(0, TOOL)
import bhayankar_run as run  # noqa: E402

FIX = os.path.join(HERE, 'fixtures')


class Workspace(unittest.TestCase):
    """A throw-away input\\ and output\\ that the runner is pointed at."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='bhayankar_wf_')
        self._old = (run.INPUT_DIR, run.OUTPUT_DIR)
        run.INPUT_DIR, run.OUTPUT_DIR = os.path.join(self.tmp, 'input'), os.path.join(self.tmp, 'output')
        os.makedirs(run.INPUT_DIR)

    def tearDown(self):
        run.INPUT_DIR, run.OUTPUT_DIR = self._old
        shutil.rmtree(self.tmp, ignore_errors=True)

    def put(self, fixture, name=None):
        shutil.copy(os.path.join(FIX, fixture), os.path.join(run.INPUT_DIR, name or fixture))

    def decisions(self, wb, workflow):
        with open(os.path.join(run.INPUT_DIR, 'decisions.json'), 'w', encoding='utf-8') as f:
            json.dump({wb: {'workflow': workflow}}, f)


class Runner(Workspace):
    def test_scan_states(self):
        self.assertEqual(run.scan(), [])
        self.put('mini.twb')
        self.assertEqual([r['state'] for r in run.scan()], ['new'])
        run.run_one('mini', build=False)
        self.assertEqual(run.scan()[0]['state'], 'blocked')  # published data source, no .tds
        self.put('orders_source.tds')  # the file set changed: runnable again
        self.assertEqual(run.scan()[0]['state'], 'changed')
        run.run_one('mini', build=False)
        self.assertEqual(run.scan()[0]['state'], 'done')
        with open(os.path.join(run.INPUT_DIR, 'mini.twb'), 'a', encoding='utf-8') as f:
            f.write('\n<!-- edited -->')
        self.assertEqual(run.scan()[0]['state'], 'changed')

    def test_published_source_blocks_until_tds_or_approval(self):
        self.put('mini.twb')
        st = run.run_one('mini', build=False)
        self.assertEqual(st['verdict'], 'BLOCKED')
        self.assertIn('Published data source', st['blockers'][0])
        self.assertNotIn('build', st['stages'])
        self.decisions('mini', {'accept_guessed_joins': True})
        self.assertEqual(run.run_one('mini', build=False)['verdict'], 'READY TO BUILD')

    def test_tds_clears_the_published_source_blocker(self):
        self.put('mini.twb')
        self.put('orders_source.tds')
        st = run.run_one('mini', build=False)
        self.assertEqual(st['blockers'], [])
        self.assertEqual(st['verdict'], 'READY TO BUILD')
        self.assertTrue(os.path.exists(os.path.join(run.out_dir('mini'), run.STATUS_MD)))

    def test_sample_data_folder_is_not_a_blocker_but_is_a_check(self):
        self.put('mini.twb')
        os.makedirs(os.path.join(run.INPUT_DIR, 'mini_Data'))
        with open(os.path.join(run.INPUT_DIR, 'mini_Data', 't.csv'), 'w') as f:
            f.write('a\n1\n')
        st = run.run_one('mini', build=False)
        self.assertEqual(st['blockers'], [])
        self.assertIn('PUBLISHED_DS_SAMPLE', [c['code'] for c in st['compare']])

    def test_security_logic_needs_an_architect(self):
        self.put('collision.twb')
        st = run.run_one('collision', build=False)
        self.assertEqual(st['verdict'], 'BLOCKED')
        self.assertTrue(any('security' in b.lower() for b in st['blockers']))
        self.decisions('collision', {'rls_approved_by': 'A. Architect'})
        self.assertFalse(any('security' in b.lower() for b in run.run_one('collision', build=False)['blockers']))

    def test_unchanged_workbook_is_skipped_unless_forced(self):
        self.put('mini.twb')
        self.put('orders_source.tds')
        run.run_one('mini', build=False)
        self.assertIn('skipped', run.run_one('mini', build=False))
        self.assertNotIn('skipped', run.run_one('mini', build=False, force=True))

    def test_status_page_is_written_and_says_numbers_are_not_checked(self):
        self.put('mini.twb')
        self.put('orders_source.tds')
        run.run_one('mini', build=False)
        text = run.read_text(os.path.join(run.out_dir('mini'), run.STATUS_MD))
        self.assertIn('Number check vs Tableau', text)
        self.assertIn('a person does this', text)

    def test_damaged_workbook_fails_without_crashing(self):
        with open(os.path.join(run.INPUT_DIR, 'bad.twb'), 'w') as f:
            f.write('not xml at all')
        st = run.run_one('bad', build=False)
        self.assertEqual(st['verdict'], 'FAILED')

    def test_unknown_workbook(self):
        self.assertEqual(run.run_one('nope')['verdict'], 'NOT FOUND')

    def test_features_workbook_lists_checks_and_design_items(self):
        self.put('features.twb')
        st = run.run_one('features', build=False)
        codes = {c['code'] for c in st['compare']}
        self.assertIn('JOINED_GRAIN', codes)
        self.assertIn('UNHANDLED_XML', codes)
        self.assertIn('PARAMETER_ACTION', {d['code'] for d in st['design']})

    def test_verdict_logic(self):
        base = {'workbook': 'x', 'blockers': [], 'stages': {'extract': {'ok': True}}}
        self.assertEqual(run.verdict(base)[0], 'READY TO BUILD')
        ok_build = {**base, 'stages': {'extract': {'ok': True}, 'build': {'ok': True}}}
        self.assertEqual(run.verdict(ok_build)[0], 'BUILT - NEEDS HUMAN CHECKS')  # never "done"
        bad = {**base, 'stages': {'extract': {'ok': True}, 'build': {'ok': False, 'error': 'x'}}}
        self.assertEqual(run.verdict(bad)[0], 'BUILD FAILED')
        tf = {**base, 'stages': {'extract': {'ok': True}, 'build': {'ok': True}, 'test': {'ok': False}}}
        self.assertEqual(run.verdict(tf)[0], 'TESTS FAILED')


if __name__ == '__main__':
    unittest.main()
