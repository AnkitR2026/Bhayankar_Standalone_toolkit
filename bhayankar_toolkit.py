r"""Bhayankar Standalone Toolkit: the name, version and folders every tool in this folder shares.

The Standalone Toolkit is the offline edition of Bhayankar: Tableau workbook in, Power BI project out, with no
internet, no Node.js and no AI. It is a separate product from the Kiro edition (Bhayankar-Toolkit-Repo); see
README.md and docs/PROJECT_LOG.md for what differs.

Folders (each can be moved with an environment variable, e.g. to work on a shared drive):
  input\       BHAYANKAR_INPUT    Tableau workbooks, .tds, sources.json, decisions.json, table_map.json
  output\      BHAYANKAR_OUTPUT   one folder per workbook: digest, rebuild spec, Power BI project, logs
  test_input\  BHAYANKAR_TEST_INPUT, test_output\  BHAYANKAR_TEST_OUTPUT   for the report tester
Python 3 standard library only.
"""
import os

NAME = 'Bhayankar Standalone Toolkit'
VERSION = '1.5.0'
RELEASED = '2026-11-07'
HERE = os.path.dirname(os.path.abspath(__file__))
RESOURCES = os.path.join(HERE, 'resources')


def _folder(env, default):
    return os.path.abspath(os.environ.get(env) or os.path.join(HERE, default))


def input_dir():
    return _folder('BHAYANKAR_INPUT', 'input')


def output_dir():
    return _folder('BHAYANKAR_OUTPUT', 'output')


def test_input_dir():
    return _folder('BHAYANKAR_TEST_INPUT', 'test_input')


def test_output_dir():
    return _folder('BHAYANKAR_TEST_OUTPUT', 'test_output')


def banner():
    return f'{NAME} {VERSION}'
