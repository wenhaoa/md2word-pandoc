import importlib.util
from pathlib import Path

script_path = Path(__file__).with_name('merge_cover.py')
spec = importlib.util.spec_from_file_location('merge_cover', script_path)
merge_cover = importlib.util.module_from_spec(spec)
spec.loader.exec_module(merge_cover)


class FakeRun:
    def __init__(self, text):
        self.text = text


class FakeParagraph:
    def __init__(self, runs):
        self.runs = [FakeRun(text) for text in runs]

    @property
    def text(self):
        return ''.join(run.text for run in self.runs)


class FakeDoc:
    def __init__(self):
        self.paragraphs = [
            FakeParagraph(['{{TI', 'TLE}}']),
            FakeParagraph(['{{DATE', '_CN}}']),
        ]


doc = FakeDoc()
merge_cover.replace_cover_placeholders(doc, {
    '{{TITLE}}': '示例技术报告',
    '{{DATE_CN}}': '2026 年 6 月',
})

assert doc.paragraphs[0].text == '示例技术报告'
assert doc.paragraphs[1].text == '2026 年 6 月'
assert doc.paragraphs[1].runs[0].text == '2026 年 6 月'
assert doc.paragraphs[1].runs[1].text == ''

print('merge_cover placeholder tests passed')
