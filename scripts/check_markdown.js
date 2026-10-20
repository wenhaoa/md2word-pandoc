const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { restoreImageCaptions } = require('./markdown_compat');

function plain(value) {
    if (Array.isArray(value)) return value.map(plain).join('');
    if (!value || typeof value !== 'object') return '';
    if (value.t === 'Str') return value.c;
    if (['Space', 'SoftBreak', 'LineBreak'].includes(value.t)) return ' ';
    if (['Code', 'CodeBlock', 'Math', 'Image', 'Link'].includes(value.t)) return '';
    return plain(value.c);
}

function check(source) {
    source = path.resolve(source);
    const ast = restoreImageCaptions(JSON.parse(execFileSync('pandoc', [source, '--from',
        'markdown-smart-subscript-superscript-implicit_figures+tex_math_single_backslash', '--to', 'json'],
        { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 })));
    const issues = [], captions = { '图': [], '表': [] }, references = [], equationTags = [], equations = [];
    // Visible captions are standalone body paragraphs. Cell text can cite a
    // figure or table but must not create a second caption with that number.
    const bodyParagraphs = new Set(ast.blocks.filter(block => ['Para', 'Plain'].includes(block.t)));
    const add = (kind, text) => issues.push({ kind, text });
    function walk(value) {
        if (Array.isArray(value)) { value.forEach(walk); return; }
        if (!value || typeof value !== 'object') return;
        if (['CodeBlock', 'Code'].includes(value.t)) {
            if (value.t === 'CodeBlock' && value.c[0][1].includes('mermaid')) add('Mermaid', 'Mermaid应先生成图片');
            return;
        }
        if (value.t === 'Math') {
            if (value.c[0].t === 'DisplayMath') {
                const tags = [...value.c[1].matchAll(/\\tag\{([^}]*)\}/g)].map(match => match[1]);
                if (tags.length !== 1 || !/^\d+-\d+$/.test(tags[0])) add('公式编号', '显示公式应有一个编号：\\tag{章-序}');
                equationTags.push(...tags);
            }
            return;
        }
        if (value.t === 'Image' || value.t === 'Link') {
            const destination = value.c[2][0];
            if (destination && !destination.startsWith('#') && (!/^[A-Za-z][\w+.-]*:/.test(destination) || /^[A-Za-z]:[/\\]/.test(destination))) {
                const target = decodeURIComponent(destination.split('#')[0].split('?')[0]);
                if (!fs.existsSync(path.resolve(path.dirname(source), target))) add('本地链接不存在', destination);
            }
        }
        if (['Para', 'Plain'].includes(value.t)) {
            const text = plain(value.c).trim();
            const caption = bodyParagraphs.has(value) && text.match(/^(图|表)\s*(\d+(?:[-.]\d+)?)\s*(\D.*)$/);
            if (caption) captions[caption[1]].push(caption[2].replace('.', '-'));
            else {
                for (const match of text.matchAll(/(图|表)(\d+(?:[-.]\d+)?)/g)) references.push(match);
                for (const match of text.matchAll(/(?:公式|式)\s*[（(](\d+-\d+)[）)]/g)) equations.push(match[1]);
            }
        }
        Object.values(value).forEach(walk);
    }
    walk(ast.blocks);
    function sequence(kind, values) {
        const counts = new Map();
        for (const value of values) {
            const split = value.split('-');
            const group = split.length === 1 ? 'global' : split[0];
            const actual = Number(split[split.length - 1]);
            const expected = (counts.get(group) || 0) + 1;
            counts.set(group, expected);
            if (actual !== expected) add(kind + '编号顺序', `${value}，本组期望${expected}`);
        }
    }
    sequence('图', captions['图']); sequence('表', captions['表']); sequence('公式', equationTags);
    for (const match of references) if (!captions[match[1]].includes(match[2].replace('.', '-'))) add('图表引用悬空', match[0]);
    for (const key of equations) if (!equationTags.includes(key)) add('公式引用悬空', key);
    // Legacy image alt captions have already been materialized above.
    for (let i = 0; i < ast.blocks.length; i++) {
        const block = ast.blocks[i];
        if (!['Para', 'Plain', 'Table'].includes(block.t)) continue;
        const hasImage = JSON.stringify(block).includes('"t":"Image"');
        if (block.t === 'Table' || hasImage) {
            const kind = block.t === 'Table' ? '表' : '图';
            const neighbours = [ast.blocks[i - 1], ast.blocks[i + 1]].filter(Boolean);
            if (!neighbours.some(item => new RegExp('^' + kind + '\\s*\\d+(?:[-.]\\d+)?\\s*\\D').test(plain(item.c).trim()))) {
                add('缺少题注', kind + '题注应独立成段；图片也可在alt中填写图号及标题');
            }
        }
    }
    return issues;
}

if (require.main === module) {
    try {
        const source = process.argv.slice(2).find(arg => !arg.startsWith('--'));
        if (!source) throw new Error('用法：node check_markdown.js 文件.md [--json]');
        const issues = check(source);
        if (process.argv.includes('--json')) console.log(JSON.stringify({ file: path.resolve(source), issues }, null, 2));
        else if (!issues.length) console.log('Markdown预检通过；标题可使用现有层级，明确编号引用予以保留。');
        else issues.forEach(issue => console.log(issue.kind + '：' + issue.text));
        process.exitCode = issues.length ? 1 : 0;
    } catch (error) { console.error(error.message); process.exitCode = 2; }
}
module.exports = { check };
