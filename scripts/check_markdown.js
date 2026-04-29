const fs = require('fs');
const path = require('path');

const mdFileInput = process.argv.slice(2).find(a => !a.startsWith('--'));

if (!mdFileInput) {
    console.error('错误：请提供 Markdown 文件路径');
    console.error('用法: node check_markdown.js <源文件.md>');
    process.exit(2);
}

const mdFile = path.resolve(mdFileInput);
if (!fs.existsSync(mdFile)) {
    console.error(`错误：文件不存在: ${mdFile}`);
    process.exit(2);
}

const content = fs.readFileSync(mdFile, 'utf8').replace(/^\uFEFF/, '');
const lines = content.split(/\r?\n/);
const issues = [];

function addIssue(lineNo, type, current, suggestion) {
    issues.push({
        lineNo,
        type,
        current: (current || '').trim().slice(0, 120),
        suggestion,
    });
}

function isBlank(index) {
    return index < 0 || index >= lines.length || lines[index].trim() === '';
}

function visibleWidth(text) {
    let width = 0;
    for (const ch of text) {
        width += /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(ch) ? 2 : 1;
    }
    return width;
}

function parsePipeCells(line) {
    return line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(cell => cell.trim());
}

function dashCount(cell) {
    return (cell.match(/-/g) || []).length;
}

function isStructuralLine(line) {
    const trimmed = line.trim();
    return trimmed === '' ||
        /^---$/.test(trimmed) ||
        /^#{1,6}\s+/.test(trimmed) ||
        /^!\[.*?\]\(.*?\)\s*$/.test(trimmed) ||
        /^```/.test(trimmed) ||
        /^\$\$/.test(trimmed) ||
        /^\|.*\|\s*$/.test(trimmed) ||
        /^>\s*/.test(trimmed) ||
        /^\s*(?:[-+*]\s+|\d+\.\s+)/.test(line) ||
        /^表(?:[0-9]+|[A-Z])-[0-9]+\s+.+/.test(trimmed);
}

function parseFrontmatter() {
    if (lines[0] !== '---') {
        return { exists: false, endIndex: -1, title: '' };
    }
    for (let i = 1; i < lines.length; i++) {
        if (lines[i] === '---') {
            const block = lines.slice(1, i);
            const titleLine = block.find(line => /^title\s*:\s*.+/.test(line));
            const title = titleLine ? titleLine.replace(/^title\s*:\s*/, '').trim().replace(/^['"]|['"]$/g, '') : '';
            return { exists: true, endIndex: i, title };
        }
    }
    return { exists: true, endIndex: -1, title: '' };
}

function nearestChapter(lineIndex) {
    for (let i = lineIndex; i >= 0; i--) {
        const m = lines[i].match(/^##\s+(?:([0-9]+)\.|附录\s+([A-Z]))/);
        if (m) return m[1] || m[2];
    }
    return null;
}

const frontmatter = parseFrontmatter();
if (!frontmatter.exists) {
    addIssue(1, 'frontmatter', lines[0], '文件开头添加 YAML frontmatter，并填写 title 字段');
} else if (frontmatter.endIndex < 0) {
    addIssue(1, 'frontmatter', '---', '补齐 frontmatter 结束行 ---');
} else if (!frontmatter.title) {
    addIssue(1, 'frontmatter', 'title', '填写非空 title，用于 Word 封面标题');
}

for (let i = Math.max(frontmatter.endIndex + 1, 0); i < lines.length; i++) {
    const line = lines[i];
    const lineNo = i + 1;

    if (/\s+$/.test(line) && line.trim() !== '') {
        addIssue(lineNo, '行尾空格', line, '删除行尾多余空格');
    }

    if (line !== '' && line.trim() === '') {
        addIssue(lineNo, '空白行含空格', line, '空白行不要保留空格或制表符');
    }

    if (line.trim() === '' && i > 0 && lines[i - 1].trim() === '') {
        addIssue(lineNo, '连续空行', line, '正式报告中连续空行压缩为一个空行');
    }

    if (!isStructuralLine(line) && !isStructuralLine(lines[i + 1] || '')) {
        addIssue(lineNo + 1, '正文硬换行', lines[i + 1], '同一段正文不要手动换行；需要分段时中间保留一个空行');
    }

    if (line.trim() === '---') {
        addIssue(lineNo, '水平线', line, '转 Word 文档中避免使用 --- 水平分隔线');
    }

    if (/^#\s+/.test(line)) {
        addIssue(lineNo, '标题层级', line, '正文从 ## 开始，封面标题使用 frontmatter title');
    }

    if (/^##\s+/.test(line) && !/^##\s+(?:[0-9]+\.\s+|附录\s+[A-Z]\s+)/.test(line)) {
        addIssue(lineNo, '一级标题格式', line, '使用 ## N. 标题 或 ## 附录 X 标题');
    }

    if (/^###\s+/.test(line) && !/^###\s+[0-9]+\.[0-9]+\s+/.test(line)) {
        addIssue(lineNo, '二级标题格式', line, '使用 ### N.N 标题');
    }

    if (/^####\s+/.test(line) && !/^####\s+[0-9]+\.[0-9]+\.[0-9]+\s+/.test(line)) {
        addIssue(lineNo, '三级标题格式', line, '使用 #### N.N.N 标题');
    }

    if (/^#{2,4}\s+/.test(line)) {
        if (!isBlank(i - 1)) {
            addIssue(lineNo, '标题前空行', line, '标题前保留一个空行，避免标题与上段正文粘连');
        }
        if (!isBlank(i + 1)) {
            addIssue(lineNo, '标题后空行', line, '标题后保留一个空行，再开始正文或图表');
        }
    }

    if (/^>\s*\[!(NOTE|WARNING|TIP|IMPORTANT|CAUTION)\]/i.test(line)) {
        addIssue(lineNo, 'GitHub提示块', line, '改为“注：”前缀的普通段落');
    }

    if (/^```mermaid\s*$/i.test(line)) {
        addIssue(lineNo, 'Mermaid代码块', line, '先转换为 PNG，再用图片引用');
    }

    if (/^\s*-\s+/.test(line)) {
        addIssue(lineNo, '无序列表', line, '正式报告优先用段落或（1）（2）（3）手动编号');
    }

    if (/\$\s*\d+(?:\.\d+)?\s*(?:\\,)?\s*\\text\{[A-Za-z%°μ]+\}\s*\$/.test(line)) {
        addIssue(lineNo, '简单单位公式', line, '简单数字+单位直接写纯文本，如 500km');
    }

    const imageMatch = line.match(/^!\[(.*?)\]\((.*?)\)\s*$/);
    if (imageMatch) {
        const caption = imageMatch[1];
        if (!/^图(?:[0-9]+|[A-Z])-[0-9]+\s+.+/.test(caption)) {
            addIssue(lineNo, '图片题注', line, '使用 ![图N-M 标题](path) 或 ![图A-M 标题](path)');
        }
        if (caption.length > 60) {
            addIssue(lineNo, '图片题注过长', caption, '题注控制在 60 字符以内');
        }
        if (!isBlank(i + 1)) {
            addIssue(lineNo, '图片后空行', line, '图片行后保留一个空行');
        }
    }

    if (/^\|.*\|\s*$/.test(line) && /^\|[\s\-:|]+\|\s*$/.test(lines[i + 1] || '')) {
        let j = i - 1;
        while (j >= 0 && lines[j].trim() === '') j--;
        const captionLine = j >= 0 ? lines[j].trim() : '';
        if (!/^表(?:[0-9]+|[A-Z])-[0-9]+\s+.+/.test(captionLine)) {
            addIssue(lineNo, '表格题注', line, '表格上方添加独立题注行：表N-M 标题');
        }
        if (i - j - 1 !== 1) {
            addIssue(lineNo, '题注表格间距', line, '表格题注与表格之间保留且仅保留一个空行');
        }
        if (!isBlank(j - 1)) {
            addIssue(j + 1, '表格前空行', captionLine, '表格题注前保留一个空行');
        }
        const headers = parsePipeCells(line);
        const separators = parsePipeCells(lines[i + 1] || '');
        const firstColumnDashMin = Math.max(6, Math.min(12, visibleWidth(headers[0] || '')));
        if (dashCount(separators[0] || '') < firstColumnDashMin) {
            addIssue(lineNo + 1, '表格列宽控制', lines[i + 1], `第一列分隔横线建议不少于 ${firstColumnDashMin} 个 -，用横线数量控制 Word 表格列宽`);
        }
        if (separators.length > 1 && separators.every(cell => dashCount(cell) <= 3)) {
            addIssue(lineNo + 1, '表格列宽控制', lines[i + 1], '不要所有列都只写 ---；按内容长短增加横线数量，尤其控制第一列宽度');
        }
    }

    const refs = line.match(/[图表](?:[0-9]+|[A-Z])-[0-9]+/g) || [];
    if (refs.length > 0 && !/^!\[图/.test(line) && !/^表(?:[0-9]+|[A-Z])-[0-9]+\s+/.test(line.trim())) {
        addIssue(lineNo, '正文图表编号引用', line, '正文建议写“如图所示”或“如下表所示”，避免固定编号和交叉引用');
    }
}

function checkSequence(prefix, regex) {
    const byChapter = new Map();
    lines.forEach((line, index) => {
        const match = line.match(regex);
        if (!match) return;
        const chapter = match[1];
        const number = Number(match[2]);
        if (!byChapter.has(chapter)) byChapter.set(chapter, []);
        byChapter.get(chapter).push({ number, lineNo: index + 1, line });
    });
    for (const [chapter, items] of byChapter.entries()) {
        const sorted = [...items].sort((a, b) => a.number - b.number);
        sorted.forEach((item, idx) => {
            const expected = idx + 1;
            if (item.number !== expected) {
                addIssue(item.lineNo, `${prefix}编号连续性`, item.line, `${chapter}章内${prefix}编号应从 1 开始连续，当前期望 ${prefix}${chapter}-${expected}`);
            }
        });
    }
}

checkSequence('图', /^!\[图([0-9]+|[A-Z])-([0-9]+)\s+/);
checkSequence('表', /^表([0-9]+|[A-Z])-([0-9]+)\s+/);

console.log(`检查文件: ${mdFile}`);
if (issues.length === 0) {
    console.log('未发现 md2word 预检问题。');
    process.exit(0);
}

console.log('');
console.log('| 行号 | 问题类型 | 当前内容 | 建议修改 |');
console.log('| ---- | -------- | -------- | -------- |');
for (const issue of issues) {
    const current = issue.current.replace(/\|/g, '\\|');
    const suggestion = issue.suggestion.replace(/\|/g, '\\|');
    console.log(`| ${issue.lineNo} | ${issue.type} | ${current} | ${suggestion} |`);
}

console.log('');
console.log(`共发现 ${issues.length} 个问题。`);
process.exit(1);
