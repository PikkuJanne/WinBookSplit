"""Fail-closed pure guard for actual M4-T01 fidelity receipts."""
from hashlib import sha256
import importlib.util
import json
import math
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[2]
HEX = re.compile(r'[0-9a-f]{64}\Z')
APPLICATION = ('WinBookSplit.ps1', 'WinBookSplit.bat', 'Export-WinBookSplitDiagnostics.ps1', 'requirements.txt',
    'engine/WinBookSplit.Runtime.ps1', 'engine/WinBookSplit.Process.ps1', 'engine/WinBookSplit.Diagnostics.ps1',
    'engine/WinBookSplit.Paths.ps1', 'engine/WinBookSplit.Logging.ps1', 'engine/WinBookSplit.Support.ps1',
    'engine/WinBookSplit.Outcomes.json', 'engine/winbooksplit_engine.py', 'engine/winbooksplit_conversion.py',
    'engine/winbooksplit_windows.py', 'engine/winbooksplit_job.py')
MATRIX = [('rich', 'manual'), ('rich', '1'), ('rich', '2'), ('scanned', 'manual')]
RANGES = {'manual': [[0, 3], [3, 6]], '1': [[0, 3], [3, 6]], '2': [[0, 1], [1, 3], [3, 4], [4, 6]]}
AUTHOR = 'Authored Å café 日本 source author'
TITLE = 'Original Å café 日本 page fidelity fixture'
NAVIGATION_CODES = {'cross_chapter_link_dropped', 'navigation_link_dropped', 'article_navigation_dropped'}
ANNOTATION_CODES = {'annotation_dropped', 'annotation_relation_dropped'}
METADATA_CODES = {'metadata_omitted', 'metadata_normalized'}


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


console = load('wbs_fidelity_console_guard', ROOT / 'tests/manual/current_launchers.py')
support = load('wbs_fidelity_support_guard', ROOT / 'tests/support/validate_support_report.py')


def need(condition, message):
    if not condition:
        raise ValueError(message)


def valid_hash(value):
    return isinstance(value, str) and HEX.fullmatch(value) is not None


def strict_json(text):
    return support.strict_json(text)


def records(text, prefix):
    return [strict_json(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]


def validate_process(process, command=None):
    need(isinstance(process, dict) and process.get('actual_process') is True and type(process.get('exit_code')) is int
         and process['exit_code'] == 0 and process.get('timed_out') is False and process.get('stdin') == 'DEVNULL'
         and type(process.get('pid')) is int and process['pid'] > 0, 'Actual finite native process evidence missing')
    need(type(process.get('elapsed_seconds')) in (int, float) and math.isfinite(process['elapsed_seconds'])
         and 0 <= process['elapsed_seconds'] < process.get('timeout_seconds', 0), 'Native elapsed/deadline evidence invalid')
    need(isinstance(process.get('command'), list) and all(isinstance(value, str) for value in process['command']), 'Exact native argv missing')
    if command is not None:
        need(process['command'] == command, 'Actual native argv disagrees with tested renderer/entry point')
    for stream in ('stdout', 'stderr'):
        need(isinstance(process.get(stream), str) and sha256(process[stream].encode('utf-8')).hexdigest() == process.get(stream + '_sha256'), 'Native stream raw hash differs')


def validate_png(page):
    need(isinstance(page, dict) and all(type(page.get(key)) is int and page[key] > 0 for key in ('page', 'bytes', 'width', 'height'))
         and valid_hash(page.get('png_sha256')) and valid_hash(page.get('pixels_sha256'))
         and isinstance(page.get('path'), str) and Path(page['path']).is_absolute()
         and Path(page['path']).name == page.get('filename'), 'Retained actual PNG dimensions/hash/path evidence invalid')


def validate_transport(transports, stdout, log_text):
    need(isinstance(transports, list) and len(transports) == 1
         and all(transports[0].get(field) is True for field in ('JobAssigned', 'ParentStopped', 'DescendantsStopped', 'StreamsComplete'))
         and type(transports[0].get('ExitCode')) is int and transports[0]['ExitCode'] == 0
         and type(transports[0].get('ResultRecordCount')) is int and transports[0]['ResultRecordCount'] == 1,
         'Actual native job/stream/no-prompt process evidence missing')
    # Ordinary Run does not log interactive-session counters. Do not invent them.
    for field in ('InteractionCount', 'ReplyCount', 'QueuedReplyCount'):
        need(field not in transports[0] or type(transports[0][field]) is int and transports[0][field] == 0,
             'NonInteractive transport recorded interaction activity')
    need(not any(marker in stdout or marker in log_text for marker in ('[WBS-INTERACTION] ', '[INTERACTION] ', '[REPLY] ')),
         'NonInteractive transport emitted an interaction request or reply')


def validate_export_process(row, process):
    directory = Path(row['input_path']).parent.parent
    manifest = Path(row['output_base']) / ('.WinBookSplit-console-' + row['console_evidence']['owner']['run_id']) / 'WinBookSplit_Run.json'
    command = [row['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
               str(directory / 'a/Export-WinBookSplitDiagnostics.ps1'), '-ManifestPath', str(manifest),
               '-OutputPath', str(Path(row['cwd']) / 'redacted-summary.json')]
    validate_process(process, command)
    need(process.get('cwd') == row['cwd'] and Path(row['cwd']) == directory / 'c',
         'Actual exporter cwd differs from the owned application control')


def validate_export_frame(stdout):
    frames = records(stdout, '[SUPPORT-EXPORT] ')
    need(len(frames) == 1 and isinstance(frames[0], dict)
         and set(frames[0]) == {'protocol', 'version', 'status', 'code', 'exit_code', 'written_count'}
         and frames[0]['protocol'] == 'winbooksplit.support-export'
         and frames[0]['status'] == 'success' and frames[0]['code'] == 'support_export_complete'
         and all(type(frames[0][key]) is int for key in ('version', 'exit_code', 'written_count'))
         and (frames[0]['version'], frames[0]['exit_code'], frames[0]['written_count']) == (1, 0, 1),
         'Authoritative actual export native/frame mismatch')


def compare_render(source, actual, selected):
    need(isinstance(source, dict) and isinstance(actual, dict) and set(source) == set(actual) == {'Poppler', 'PDFium'}, 'Both actual renderer comparisons required')
    for name in source:
        validate_process(source[name]['process'])
        validate_process(actual[name]['process'])
        expected = source[name]['pages'][selected[0]:selected[1]]
        observed = actual[name]['pages']
        need(len(expected) == len(observed) == selected[1] - selected[0], 'Actual output render page count differs')
        for index, (before, after) in enumerate(zip(expected, observed), 1):
            validate_png(before)
            validate_png(after)
            need(after['page'] == index and all(before[key] == after[key] for key in ('width', 'height', 'png_sha256', 'pixels_sha256')), 'Same-renderer source/output pixels or PNG bytes differ')


def validate_render_binding(rendered, input_path, directory, prefix, renderer, secondary_python):
    directory = Path(directory)
    need(directory.is_absolute(), 'Persistent render path must be absolute')
    need(set(rendered) == {'Poppler', 'PDFium'}, 'Both renderer route bindings required')
    primary_prefix = directory / (prefix + '-poppler')
    secondary_prefix = directory / (prefix + '-pdfium')
    validate_process(rendered['Poppler']['process'], [renderer, '-r', '72', '-cropbox', '-png', input_path, str(primary_prefix)])
    validate_process(rendered['PDFium']['process'], [secondary_python, '-I', '-B', str(ROOT / 'tests/fidelity/render_pdfium.py'),
                                                   '--input', input_path, '--output-prefix', str(secondary_prefix)])
    for name, expected_prefix in (('Poppler', primary_prefix), ('PDFium', secondary_prefix)):
        need(rendered[name]['process']['cwd'] == str(directory), 'Renderer cwd differs from retained exclusive artifact directory')
        for index, page in enumerate(rendered[name]['pages'], 1):
            validate_png(page)
            match = re.fullmatch(re.escape(expected_prefix.name) + r'-([0-9]+)\.png', page['filename'])
            need(match is not None and int(match[1]) == page['page'] == index and Path(page['path']).parent == directory,
                 'Retained renderer page escaped exact prefix/path/order')


def validate_output(row, output, entry, expected):
    start, end = entry['start'], entry['end']
    count = end - start
    need(output.get('filename') == entry['filename'] and output.get('start') == start and output.get('end') == end
         and output.get('sha256') == entry['sha256'] and type(output.get('bytes')) is int
         and output['bytes'] == entry['size_bytes'] and output['bytes'] > 0, 'Actual output filename/range/byte hash differs from execution')
    pages = output.get('pages')
    need(isinstance(pages, list) and len(pages) == count and output.get('expected_pages') == expected, 'Reopened page observation missing')
    for local, (actual, source) in enumerate(zip(pages, expected)):
        need(set(actual) == set(source) == {'page_id', 'text', 'content_sha256', 'images', 'geometry', 'static_annotations', 'links'}, 'Exact structural page snapshot shape changed')
        need(type(actual['page_id']) is int and actual['page_id'] == start + local + 1, 'Physical first/last/page identity differs')
        need(all(actual[key] == source[key] for key in ('text', 'content_sha256', 'images', 'geometry', 'static_annotations')), 'Text/vector/image/geometry/static appearance changed')
        expected_links = []
        for link in source['links']:
            target = link['target_page']
            if type(target) is int and start <= target < end:
                expected_links.append({**link, 'target_page': target - start, 'form': 'destination'})
        need(isinstance(actual['links'], list) and all(
            isinstance(link, dict) and set(link) == {'name', 'target_page', 'form', 'fit', 'args', 'rect', 'border'}
            and isinstance(link['name'], str) and type(link['target_page']) is int and 0 <= link['target_page'] < count
            and link['form'] == 'destination' and link['fit'] in {'/XYZ', '/Fit', '/FitH', '/FitV', '/FitR', '/FitB', '/FitBH', '/FitBV'}
            and isinstance(link['args'], list) and all(value is None or type(value) in (int, float) and math.isfinite(value) for value in link['args'])
            and isinstance(link['rect'], list) and len(link['rect']) == 4
            and isinstance(link['border'], list) and len(link['border']) == 3
            and all(type(value) in (int, float) and math.isfinite(value) for value in link['rect'] + link['border'])
            for link in actual['links']), 'Rebased destination fit/coordinates use invalid or Boolean numeric aliases')
        need(actual['links'] == expected_links, 'Intra-segment links not exactly remapped or invalid/cross-segment link survived')
        geometry = actual['geometry']
        need(isinstance(geometry, dict) and type(geometry.get('rotation')) is int and geometry['rotation'] in {0, 90, 180, 270}
             and all(isinstance(geometry.get(key), list) and len(geometry[key]) == 4
                     and all(type(value) in (int, float) and math.isfinite(value) for value in geometry[key])
                     for key in ('media', 'crop', 'trim', 'bleed', 'art')), 'Reopened geometry uses invalid/Boolean numeric aliases')
    inventory = output.get('inventory')
    images = sorted({image['decoded_sha256'] for page in expected for image in page['images']})
    marker_ids = list(range(start + 1, end + 1)) if row['kind'] == 'rich' else []
    need(isinstance(inventory, dict) and type(inventory.get('all_page_object_count')) is int and inventory['all_page_object_count'] == count
         and inventory.get('serialized_content_marker_ids') == marker_ids and inventory.get('all_image_decoded_sha256') == images,
         'Serialized output contains hidden off-segment pages/content/images or omitted selected resources')
    chapter = re.sub(r'^\d+ - ', '', entry['filename'][:-4])
    metadata = output.get('metadata')
    source_title = TITLE + (' image-only' if row['kind'] == 'scanned' else '')
    need(isinstance(metadata, dict) and metadata.get('/Title') == chapter and metadata.get('/Author') == AUTHOR
         and metadata.get('/Creator') == 'WinBookSplit' and metadata.get('/Producer') == 'WinBookSplit 1.0.0-dev (pypdf 6.19.0)'
         and metadata.get('/Subject') == 'Source: ' + source_title and '/AuthoredPrivateField' not in metadata,
         'Chapter/source metadata incorrect or original private metadata copied')
    need(output.get('outline') == [{'title': chapter, 'page': 0}], 'Useful chapter-start bookmark differs from safe chapter label/page0')
    need(set(output.get('catalog_keys', [])) <= {'/Type', '/Pages', '/Outlines', '/PageMode'} and {'/Type', '/Pages', '/Outlines'} <= set(output['catalog_keys']), 'Arbitrary source catalog copied into output')
    need(isinstance(output.get('annotations'), list) and len(output['annotations']) == sum(len(page['static_annotations']) + len(page['links']) for page in pages)
         and all(item.get('P_matches_page') is True and item.get('source_relations_absent') is True and item.get('unsupported_action_absent') is True for item in output['annotations']),
         'Annotation references retain another source page/relation or unsupported action')
    compare_render(row['source_renders'], output['renders'], [start, end])


def validate_case(row, *, cleanup_complete=True):
    need(isinstance(row, dict) and (row.get('kind'), row.get('mode')) in MATRIX and row.get('host_id') in {'PS51', 'PS7'}, 'Unexpected native fidelity case')
    validate_process(row)
    need(row.get('source_read_only_observed') is True and row.get('application_unchanged') is True, 'Actual source/read-only or application guard missing')
    before = row.get('source_observations_before')
    need(isinstance(before, dict) and set(before) == {'source', 'neighbor', 'prior'} and before == row.get('source_observations_after'), 'Source/neighbor/prior identity changed during actual run')
    if cleanup_complete:
        need(row.get('passed') is True and before == row.get('source_observations_after_cleanup') and row.get('output_members_after_cleanup') == ['prior-output.pdf'], 'Owned cleanup or source identity proof missing')
    for value in before.values():
        need(isinstance(value, dict) and valid_hash(value.get('sha256')) and all(type(value.get(key)) is int and value[key] >= 0 for key in ('size_bytes', 'device', 'inode', 'attributes')), 'Captured literal-file identity invalid')
    need(isinstance(row.get('application_sha256'), dict) and set(row['application_sha256']) == set(APPLICATION)
         and all(valid_hash(value) for value in row['application_sha256'].values()), 'Actual copied application source map missing')
    final = row.get('outcome')
    need(records(row['stdout'], '[OUTCOME] ') == [final] and isinstance(final, dict) and final.get('protocol') == 'winbooksplit.outcome'
         and type(final.get('version')) is int and final['version'] == 1 and final.get('status') == 'success' and final.get('code') == 'split_complete'
         and type(final.get('exit_code')) is int and final['exit_code'] == 0 and final.get('mode') == row['mode'], 'Native/stdout authoritative successful outcome contradictory')
    result = final.get('engine_result')
    need(isinstance(result, dict) and result.get('protocol') == 'winbooksplit.result' and type(result.get('version')) is int and result['version'] == 1 and result.get('mode') == row['mode']
         and result.get('status') == 'success' and result.get('code') == 'split_complete' and result.get('exit_code') == 0
         and row.get('engine_records') == [result] and result.get('plan') == row.get('plan'), 'Sole successful terminal engine frame missing')
    plan = row['plan']
    ranges = [[0, 2], [2, 4]] if row['kind'] == 'scanned' else RANGES[row['mode']]
    pages = 4 if row['kind'] == 'scanned' else 6
    need(isinstance(plan, dict) and plan.get('total_pages') == pages and plan.get('mode') == row['mode']
         and [[entry['start'], entry['end']] for entry in plan.get('entries', [])] == ranges
         and plan.get('source_identity', {}).get('sha256') == before['source']['sha256']
         and plan.get('source_identity', {}).get('size_bytes') == before['source']['size_bytes'], 'Captured source/physical shared plan differs')
    need(isinstance(plan.get('coverage'), dict) and plan['coverage'].get('complete') is True
         and type(plan['coverage'].get('covered_pages')) is int and type(plan['coverage'].get('section_count')) is int
         and plan['coverage'] == {'complete': True, 'covered_pages': pages, 'section_count': len(ranges)}, 'Complete typed physical coverage missing')
    execution = result.get('execution')
    need(isinstance(execution, dict) and execution.get('written_count') == final.get('written_count') == result.get('written_count') == len(ranges)
         and execution.get('source_identity') == plan['source_identity'] and execution.get('mode') == row['mode'] and execution.get('total_pages') == pages
         and execution.get('coverage') == plan['coverage'] and execution.get('final_directory') == final.get('final_directory')
         and Path(execution['final_directory']).parent == Path(row['output_base']), 'Successful execution not bound to exact nonempty plan/source')
    outputs = execution.get('outputs')
    need(isinstance(outputs, list) and len(outputs) == len(ranges) and len(row.get('publication', [])) == len(outputs), 'Exact reopened publication missing')
    for entry, planned, bounds, output in zip(outputs, plan['entries'], ranges, row['publication']):
        need(all(type(entry.get(key)) is int and type(planned.get(key)) is int for key in ('start', 'end', 'sequence'))
             and all(entry.get(key) == value for key, value in planned.items()) and [entry.get('start'), entry.get('end')] == bounds
             and type(entry.get('page_count')) is int and entry['page_count'] == bounds[1] - bounds[0]
             and valid_hash(entry.get('sha256')) and type(entry.get('size_bytes')) is int and entry['size_bytes'] > 0, 'Execution changed validated plan entry/count/hash')
        validate_output(row, output, entry, row['source_pages'][bounds[0]:bounds[1]])
    need(row['publication_manifest'] == execution.get('manifest') == strict_json(row['publication_manifest_raw'])
         and sha256(row['publication_manifest_raw'].encode()).hexdigest() == row['publication_manifest_sha256'], 'Published output manifest raw/hash binding differs')
    need(set(row.get('publication_members', [])) == {'.WinBookSplit-owner.json', 'WinBookSplit_Manifest.json', *(entry['filename'] for entry in outputs)}, 'Publication owns unexpected member')
    text = row.get('console_log')
    need(isinstance(text, str) and sha256(text.encode()).hexdigest() == row.get('console_log_sha256'), 'Finalized UTF8 console raw hash differs')
    console.validate_console_evidence(row.get('console_evidence'), final, row['console_log_sha256'], source_path=row['input_path'], source_observation=before['source'], log_text=text)
    validate_transport(row.get('process_summaries'), row['stdout'], text)
    need('Done.' in row['stdout'] and 'Output: ' + execution['final_directory'] in row['stdout'], 'Truthful positive output/count summary missing')
    need(not any(marker in row['stdout'] for marker in ('Enter selection', 'Write these chapter PDFs', 'Open the completed folder', '[OPEN] ')), 'NonInteractive prompted or opened a path')
    codes = {warning.get('code') for warning in result.get('warnings', [])}
    if row['kind'] == 'rich':
        need({'cross_chapter_link_dropped', 'navigation_link_dropped'} <= codes and 'cross_chapter_link_dropped' in text and 'navigation_link_dropped' in text,
             'Cross-segment/malformed navigation dropped without visible fixed warnings')
        need(all('[WARNING] ' + code + ':' in row['stdout'] and '[WARNING] ' + code + ':' in text
                 for code in ('cross_chapter_link_dropped', 'navigation_link_dropped')),
             'Fixed navigation warning not displayed and logged as an actual categorized line')
        totals = {'cross_chapter_link_dropped': 0, 'navigation_link_dropped': 0}
        for entry in plan['entries']:
            counts = {'cross_chapter_link_dropped': 0, 'navigation_link_dropped': 0}
            for page in row['source_pages'][entry['start']:entry['end']]:
                for link in page['links']:
                    target = link['target_page']
                    if type(target) is not int:
                        counts['navigation_link_dropped'] += 1
                    elif not entry['start'] <= target < entry['end']:
                        counts['cross_chapter_link_dropped'] += 1
            for code, count in counts.items():
                matching = [warning for warning in entry['warnings'] if warning.get('code') == code]
                need(len(matching) == (1 if count else 0) and (not count or matching[0]['message'].startswith(f"Chapter {entry['sequence']}: {count}. ")),
                     'Chapter navigation warning count differs from actual omitted source links')
                totals[code] += count
        for code, count in totals.items():
            matching = [warning for warning in result['warnings'] if warning.get('code') == code]
            need(len(matching) == 1 and matching[0]['message'].startswith(f'Count: {count}. '), 'Aggregate navigation warning count differs from actual omissions')
    exported = row.get('support_export')
    if row['kind'] == 'rich' and row['mode'] == 'manual':
        need(isinstance(exported, dict), 'Actual new warned rich run export missing')
        validate_export_process(row, exported['process'])
        summary = strict_json(exported['summary_raw'])
        need(summary == exported['summary'] and sha256(exported['summary_raw'].encode()).hexdigest() == exported['summary_sha256'], 'Redacted export raw/hash binding differs')
        support.validate_summary(summary, row['console_evidence']['run_manifest'])
        validate_export_frame(exported['process']['stdout'])
        need(all(secret not in exported['summary_raw'] for secret in (AUTHOR, TITLE, row['input_path'], 'WBS-FID-PAGE-', 'MOCK_LOCAL_METADATA_NOT_FOR_OUTPUT_CATALOG')), 'Redacted warned-run summary leaked fixture/private identifiers')
    else:
        need(exported is None, 'Unexpected extra fidelity exporter control')


def validate_fidelity_report(report, requested_shells, renderer_path, secondary_python_path, *, cleanup_complete=True):
    need(isinstance(report, dict) and type(report.get('schema_version')) is int and report['schema_version'] == 1 and report.get('task_id') == 'M4-T01'
         and report.get('result') == 'FIDELITY_REGRESSION_PASSED' and report.get('success') is True and type(report.get('exit_code')) is int and report['exit_code'] == 0
         and report.get('acceptance_ids') == ['AC-066', 'AC-067', 'AC-068', 'AC-069'] and report.get('source_unchanged') is True
         and (not cleanup_complete or report.get('owned_temp_removed') is True) and report.get('human_or_GUI_tested') is False
         and report.get('render_artifacts_retained') is True, 'Fidelity aggregate/source/cleanup scope invalid')
    hosts = report.get('host_cases')
    need(isinstance(hosts, list) and len(hosts) == 2 and {host.get('id') for host in hosts} == {'PS51', 'PS7'} and len(requested_shells) == 2
         and {str(Path(host['shell_executable']).resolve()).casefold() for host in hosts} == {str(Path(path).resolve()).casefold() for path in requested_shells}, 'Actual requested two supported hosts missing')
    for host in hosts:
        need(host.get('passed') is True and host.get('exit_code') == 0 and host.get('host_major') == (5 if host['id'] == 'PS51' else 7)
             and host.get('stored_policies') == host.get('policies_after'), 'Native host/settings proof invalid')
    tested = report.get('tested_path_sha256')
    need(isinstance(tested, dict) and all(name in tested for name in APPLICATION) and all(valid_hash(value) for value in tested.values()), 'Raw tested source map missing')
    renderers = report.get('renderers')
    need(isinstance(renderers, dict) and set(renderers) == {'primary', 'secondary'}, 'Actual two-renderer probes missing')
    primary, secondary = renderers['primary'], renderers['secondary']
    need(primary.get('path') == renderer_path and primary.get('version') == '26.07.0' and valid_hash(primary.get('sha256'))
         and secondary.get('python_path') == secondary_python_path and valid_hash(secondary.get('python_sha256'))
         and secondary.get('script_sha256') == tested.get('tests/fidelity/render_pdfium.py')
         and valid_hash(secondary.get('module_sha256')) and isinstance(secondary.get('native_library_sha256'), dict)
         and secondary['native_library_sha256'] and all(valid_hash(value) for value in secondary['native_library_sha256'].values()), 'Renderer paths/source/probe hashes differ')
    validate_process(primary['probe'], [renderer_path, '-v'])
    validate_process(secondary['probe'], [secondary_python_path, '-I', '-B', str(ROOT / 'tests/fidelity/render_pdfium.py'), '--probe'])
    details = strict_json(secondary['probe']['stdout'])
    need(all(secondary.get(key) == value for key, value in details.items()), 'Actual secondary renderer version observation differs')
    provenance = report.get('fixture_provenance')
    need(isinstance(provenance, dict) and provenance.get('private_data') is False and provenance.get('generator_sha256') == tested.get('tests/fidelity/generate_fidelity_fixtures.py')
         and set(provenance.get('fixtures', {})) == {'rich', 'scanned'}, 'Authored original fixture/source provenance missing')
    cases = report.get('cases')
    need(isinstance(cases, list) and len(cases) == 8 and report.get('case_count') == 8 and type(report['case_count']) is int
         and {row.get('id') for row in cases} == {host + '-' + kind + '-' + mode for host in ('PS51', 'PS7') for kind, mode in MATRIX}, 'Complete actual fidelity matrix missing/duplicated')
    for row in cases:
        host = next(host for host in hosts if host['id'] == row['host_id'])
        fixture = provenance['fixtures'][row['kind']]
        need(row['shell_executable'] == host['shell_executable'] and row['host_version'] == host['host_version']
             and row['application_sha256'] == {name: tested[name] for name in APPLICATION}
             and row['source_observations_before']['source']['sha256'] == fixture['sha256']
             and row['source_pages'] == fixture['snapshots'] and row['source_metadata'] == fixture['metadata'], 'Actual source/host/provenance differs')
        need(row['render_directory'] == str(Path(report['render_directory']) / row['id']), 'Native case render directory escaped retained artifact root')
        validate_render_binding(row['source_renders'], row['input_path'], row['render_directory'], 'source', renderer_path, secondary_python_path)
        for number, output in enumerate(row['publication'], 1):
            path = str(Path(row['outcome']['engine_result']['execution']['final_directory']) / output['filename'])
            validate_render_binding(output['renders'], path, row['render_directory'], 'output-' + str(number), renderer_path, secondary_python_path)
        parameters = ['-InputFile', row['input_path'], '-OutputDirectory', row['output_base'], '-PythonPath', report['environment']['python_executable'],
                      '-Mode', 'Manual' if row['mode'] == 'manual' else 'Auto', '-NonInteractive']
        parameters += ['-StartPages', '1,4' if row['kind'] == 'rich' else '1,3'] if row['mode'] == 'manual' else ['-BookmarkLevel', row['mode']]
        need(row['parameters'] == parameters, 'Actual fidelity split choices differ from fixed matrix')
        validate_process(row, [host['shell_executable'], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File',
                              str(Path(row['input_path']).parent.parent / 'a/WinBookSplit.ps1'), *parameters])
        validate_case(row, cleanup_complete=cleanup_complete)
