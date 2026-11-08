r"""Bhayankar scaffold: writes the shell of a Power BI project (PBIP) with no Node.js and no internet.

Writes the same files, with the same content, as `powerbi-report-author scaffold --offline` (Microsoft
powerbi-report-authoring-cli 0.4.0, MIT), with three differences:
  - no starter page: the build writes every page itself;
  - the .platform logicalId is derived from the project name instead of being random, so two builds of the same
    workbook are byte-identical and a diff shows only real changes;
  - the .SemanticModel folder gets its .platform too (Power BI Desktop writes one when it saves a project).

  <dir>\<pbip name>.pbip
  <dir>\<item>.Report\.platform, definition.pbir
  <dir>\<item>.Report\definition\version.json, report.json, pages\pages.json
  <dir>\<item>.Report\StaticResources\SharedResources\BaseThemes\CY26SU10.json   (copied from resources\themes)
  <dir>\<item>.SemanticModel\.platform
Python 3 standard library only.
"""
import json
import os
import shutil
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
THEMES = os.path.join(HERE, 'resources', 'themes')
FABRIC = 'https://developer.microsoft.com/json-schemas/fabric'
SCHEMA = {
    'pbipProperties': f'{FABRIC}/pbip/pbipProperties/1.0.0/schema.json',
    'platformProperties': f'{FABRIC}/gitIntegration/platformProperties/2.0.0/schema.json',
    'definitionProperties': f'{FABRIC}/item/report/definitionProperties/2.0.0/schema.json',
    'versionMetadata': f'{FABRIC}/item/report/definition/versionMetadata/1.0.0/schema.json',
    'pagesMetadata': f'{FABRIC}/item/report/definition/pagesMetadata/1.1.0/schema.json',
    'report': f'{FABRIC}/item/report/definition/report/3.3.0/schema.json',
}
BASE_THEME = 'CY26SU10'
NAMESPACE = uuid.UUID('8f3b0a52-5d3c-4c8e-9d2b-6b1f2f6e7a10')  # fixed: the same name always gives the same id


def logical_id(kind, name):
    """Stable GUID for an item (.platform config.logicalId)."""
    return str(uuid.uuid5(NAMESPACE, f'{kind}:{name}'))


def write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(json.dumps(data, indent=2, ensure_ascii=False) + '\n')


def platform(kind, display_name, id_key=None):
    return {'$schema': SCHEMA['platformProperties'],
            'metadata': {'type': kind, 'displayName': display_name},
            'config': {'version': '2.0', 'logicalId': logical_id(kind, id_key or display_name)}}


def report_json():
    return {
        '$schema': SCHEMA['report'],
        'themeCollection': {'baseTheme': {'name': BASE_THEME,
                                          'reportVersionAtImport': {'visual': '2.11.0', 'report': '3.4.0', 'page': '2.3.1'},
                                          'type': 'SharedResources'}},
        'resourcePackages': [{'name': 'SharedResources', 'type': 'SharedResources',
                              'items': [{'name': BASE_THEME, 'path': f'BaseThemes/{BASE_THEME}.json', 'type': 'BaseTheme'}]}],
        'settings': {'useStylableVisualContainerHeader': True, 'exportDataMode': 'AllowSummarized',
                     'defaultDrillFilterOtherVisuals': True, 'allowChangeFilterTypes': True, 'useEnhancedTooltips': True,
                     'useDefaultAggregateDisplayName': True},
    }


def scaffold(pbip_dir, item, model_path=None, pbip_name=None):
    """Write the project shell; returns the .Report folder. item names the .Report/.SemanticModel folders, pbip_name
    the .pbip file (default: item). model_path is relative to the .Report folder (default ../<item>.SemanticModel)."""
    report_dir = os.path.join(pbip_dir, f'{item}.Report')
    model_dir = os.path.join(pbip_dir, f'{item}.SemanticModel')
    write_json(os.path.join(pbip_dir, f'{pbip_name or item}.pbip'),
               {'$schema': SCHEMA['pbipProperties'], 'version': '1.0',
                'artifacts': [{'report': {'path': f'{item}.Report'}}], 'settings': {'enableAutoRecovery': True}})
    write_json(os.path.join(report_dir, '.platform'), platform('Report', item, pbip_name or item))
    write_json(os.path.join(report_dir, 'definition.pbir'),
               {'$schema': SCHEMA['definitionProperties'], 'version': '4.0',
                'datasetReference': {'byPath': {'path': model_path or f'../{item}.SemanticModel'}}})
    write_json(os.path.join(report_dir, 'definition', 'version.json'), {'$schema': SCHEMA['versionMetadata'], 'version': '2.0.0'})
    write_json(os.path.join(report_dir, 'definition', 'report.json'), report_json())
    write_json(os.path.join(report_dir, 'definition', 'pages', 'pages.json'), {'$schema': SCHEMA['pagesMetadata'], 'pageOrder': []})
    theme_dir = os.path.join(report_dir, 'StaticResources', 'SharedResources', 'BaseThemes')
    os.makedirs(theme_dir, exist_ok=True)
    shutil.copyfile(os.path.join(THEMES, f'{BASE_THEME}.json'), os.path.join(theme_dir, f'{BASE_THEME}.json'))
    write_json(os.path.join(model_dir, '.platform'), platform('SemanticModel', item, pbip_name or item))
    return report_dir
