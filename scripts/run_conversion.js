const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { cleanSpaces, restoreImageCaptions } = require('./markdown_compat');
const { prepareTableLayouts } = require('./table_layouts');

const READER = 'markdown-smart-subscript-superscript-implicit_figures+tex_math_single_backslash';

function formatTimestamp(instant = Date.now()) {
    return new Date(instant + 8 * 3600000).toISOString().slice(0, 19)
        .replace(/[-:]/g, '').replace('T', '_');
}

function publishOutput(working, outputDir, baseName, stamp) {
    // Keep the original filename; add a numeric suffix only on a real collision.
    for (let attempt = 0; ; attempt++) {
        const suffix = attempt ? `_${attempt + 1}` : '';
        const target = path.join(outputDir, `${baseName}${stamp}${suffix}.docx`);
        try {
            fs.copyFileSync(working, target, fs.constants.COPYFILE_EXCL);
            return target;
        } catch (error) {
            if (error.code !== 'EEXIST') throw error;
        }
    }
}

function openDocument(filePath, environment = process.env, execute = execFileSync) {
    if (process.platform !== 'win32') throw new Error('--open仅适用于Windows');
    // Wait for the launcher and capture failures; do not wait for Word to close.
    execute('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command',
        '$ErrorActionPreference = "Stop"; ' +
        '$choice = (Get-ItemProperty -LiteralPath "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.docx\\UserChoice" -ErrorAction SilentlyContinue).ProgId; ' +
        '$wordPath = $null; ' +
        'if (!$choice -or $choice -like "Word.Document.*") { ' +
        'foreach ($key in @("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE", "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE", "HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE")) { ' +
        '$candidate = (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue)."(default)"; ' +
        'if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) { $wordPath = $candidate; break } } }; ' +
        'if ($wordPath) { ' +
        'Start-Process -FilePath $wordPath -ArgumentList @("/x", ([char]34 + $env:MD2WORD_OPEN_FILE + [char]34)) -ErrorAction Stop ' +
        '} else { Start-Process -FilePath $env:MD2WORD_OPEN_FILE -ErrorAction Stop }'],
        { env: { ...environment, MD2WORD_OPEN_FILE: filePath }, windowsHide: true, encoding: 'utf8' });
}

function plain(value) {
    if (Array.isArray(value)) return value.map(plain).join('');
    if (!value || typeof value !== 'object') return typeof value === 'string' ? value : '';
    if (value.t === 'Str' || value.t === 'MetaString') return value.c;
    if (['Space', 'SoftBreak', 'LineBreak'].includes(value.t)) return ' ';
    if (value.t === 'Code' || value.t === 'Math') return value.c[1];
    return plain(value.c);
}

function prepareAst(ast) {
    const plan = [];
    function visit(value) {
        if (Array.isArray(value)) { value.forEach(visit); return; }
        if (!value || typeof value !== 'object') return;
        if (value.t === 'Math' && value.c[0].t === 'DisplayMath') {
            const tags = [...value.c[1].matchAll(/\\tag\{([^}]*)\}/g)];
            if (tags.length > 1 || (tags.length && !/^\d+-\d+$/.test(tags[0][1]))) {
                throw new Error('显示公式编号无效或一个公式有多个编号');
            }
            plan.push(tags.length ? tags[0][1] : null);
            value.c[1] = value.c[1].replace(/\\tag\{[^}]*\}/g, '');
        }
        Object.values(value).forEach(visit);
    }
    // Do not treat math in metadata as a body equation.
    visit(ast.blocks);
    return plan;
}

function main(argv) {
    const flags = new Set(['--open', '--no-caption', '--citations', '--heading-mode', '--preserve-text']);
    let input, headingMode = 'auto';
    for (let i = 0; i < argv.length; i++) {
        if (argv[i] === '--heading-mode') {
            headingMode = argv[++i];
            if (!['auto', 'shift', 'chapters'].includes(headingMode)) throw new Error('--heading-mode 使用 auto、shift 或 chapters');
        } else if (argv[i].startsWith('--')) {
            if (!flags.has(argv[i])) throw new Error('未知参数：' + argv[i]);
        } else if (!input) input = argv[i];
        else throw new Error('一次只转换一个源文件');
    }
    if (!input) throw new Error('用法：node run_conversion.js 文件.md [--open] [--no-caption] [--citations] [--preserve-text] [--heading-mode auto|shift|chapters]');
    const source = path.resolve(input);
    if (!fs.statSync(source).isFile()) throw new Error('源文件不是普通文件');
    const outputDir = path.dirname(source);
    const skillDir = path.resolve(process.env.MD2WORD_SKILL_DIR || path.join(__dirname, '..'));
    const scriptDir = path.join(skillDir, 'scripts');
    const referenceDoc = path.join(skillDir, 'templates', 'md2word模板.docx');
    const filterScript = path.join(scriptDir, 'style_filter.lua');
    const steps = ['merge_cover.py', 'inline_math_to_text.py', 'merge_math_runs.py',
        'style_table_headers.py', 'autofit_tables.py', 'equation_numbering.py',
        'xref_sections.py', 'validate_docx.py', 'landscape_tables.py', 'table_sections.py'];
    if (!argv.includes('--no-caption')) steps.push('add_captions.py', 'xref_fields.py');
    if (argv.includes('--citations')) steps.push('format_citations.py');
    for (const file of [referenceDoc, filterScript, ...steps.map(name => path.join(scriptDir, name))]) {
        if (!fs.existsSync(file)) throw new Error('缺少必要文件：' + file);
    }
    const env = { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' };
    const invoke = (program, args, options = {}) => execFileSync(program, args,
        { encoding: 'utf8', env, maxBuffer: 64 * 1024 * 1024, ...options });
    invoke('pandoc', ['--version']);
    invoke('python', ['-c', 'import docx, docxcompose, lxml']);
    const tempDir = fs.mkdtempSync(path.join(outputDir, '.md2word-'));
    const working = path.join(tempDir, 'body.docx');
    const astFile = path.join(tempDir, 'source.json');
    const normalizedInput = path.join(tempDir, 'source.md');
    const planFile = path.join(tempDir, 'equations.json');
    const tablePlanFile = path.join(tempDir, 'table_layouts.json');
    const baseline = path.join(tempDir, 'baseline.json');
    const before = crypto.createHash('sha256').update(fs.readFileSync(source)).digest('hex');
    let completed = false;
    function python(name, args) {
        process.stdout.write(invoke('python', [path.join(scriptDir, name), ...args]));
    }
    try {
        console.log('源文件：' + source);
        const original = fs.readFileSync(source, 'utf8').replace(/^\uFEFF/, '');
        fs.writeFileSync(normalizedInput, argv.includes('--preserve-text') ? original : cleanSpaces(original), 'utf8');
        const ast = restoreImageCaptions(JSON.parse(invoke('pandoc', [normalizedInput, '--from', READER, '--to', 'json'])));
        const h1 = ast.blocks.filter(block => block.t === 'Header' && block.c[0] === 1);
        const isTitle = h1.length === 1 && (headingMode === 'shift' ||
            (headingMode === 'auto' && !/^\s*(?:\d+(?:\.|、|\s)|第\S*章|[一二三四五六七八九十百]+[、.．])/.test(plain(h1[0].c[2]))));
        const title = plain(ast.meta.title) || (isTitle ? plain(h1[0].c[2]) : path.parse(source).name);
        ast.meta.md2word_heading_mode = { t: 'MetaString', c: headingMode };
        const plan = prepareAst(ast);
        const tablePlan = prepareTableLayouts(ast);
        fs.writeFileSync(astFile, JSON.stringify(ast), 'utf8');
        fs.writeFileSync(planFile, JSON.stringify(plan), 'utf8');
        fs.writeFileSync(tablePlanFile, JSON.stringify(tablePlan), 'utf8');
        invoke('pandoc', [astFile, '--from', 'json', '--to', 'docx', '-o', working,
            '--reference-doc', referenceDoc, '--lua-filter', filterScript,
            '--resource-path', outputDir, '--standalone'], { stdio: 'inherit' });
        // Capture the body before cover composition and all mutating steps.
        python('validate_docx.py', [working, '--snapshot', baseline]);
        if (tablePlan.landscape.length && JSON.parse(fs.readFileSync(baseline, 'utf8')).tables !== tablePlan.table_count) {
            throw new Error('表格布局计划与Word正文表格数量不一致，已停止以避免旋转错误表格');
        }
        python('inline_math_to_text.py', [working]);
        python('merge_math_runs.py', [working]);
        if (!argv.includes('--no-caption')) python('add_captions.py', [working]);
        python('style_table_headers.py', [working]);
        python('autofit_tables.py', [working, '--body-only']);
        python('equation_numbering.py', [working, '--plan-file', planFile]);
        if (!argv.includes('--no-caption')) python('xref_fields.py', [working, '--md', source]);
        python('xref_sections.py', [working]);
        // Bracketed numbers may be array indices; citation styling is opt-in.
        if (argv.includes('--citations')) python('format_citations.py', [working]);
        python('validate_docx.py', [working, '--baseline', baseline]);
        python('merge_cover.py', [referenceDoc, working, working, '--title', title]);
        if (tablePlan.landscape.length) python('landscape_tables.py', [working, '--plan', tablePlanFile]);
        python('validate_docx.py', [working, '--baseline', baseline, '--with-cover']);
        const after = crypto.createHash('sha256').update(fs.readFileSync(source)).digest('hex');
        if (after !== before) throw new Error('转换期间源文件发生变化，未发布输出，请重新转换');
        const stamp = formatTimestamp();
        const finalOutput = publishOutput(working, outputDir, path.parse(source).name, stamp);
        completed = true;
        console.log('转换及结构检查通过。输出文件：' + finalOutput);
        console.log('打开Word后更新目录和域，并检查分页、表格宽度及公式位置。');
        if (argv.includes('--open')) {
            try {
                openDocument(finalOutput, env);
                console.log('已提交系统打开Word文档。');
            } catch (error) {
                const reason = (error.stderr || error.message).toString().trim();
                throw new Error('文档已生成，但自动打开失败：' + reason + '\n输出文件：' + finalOutput);
            }
        }
        return finalOutput;
    } finally {
        if (completed) {
            // Delete only the unique directory created by this invocation.
            if (path.dirname(tempDir) !== outputDir || !path.basename(tempDir).startsWith('.md2word-')) {
                throw new Error('临时目录路径校验失败，保留现场');
            }
            fs.rmSync(tempDir, { recursive: true, force: true });
        } else {
            console.error('转换未完成，诊断文件保留于：' + tempDir);
        }
    }
}

if (require.main === module) {
    try { main(process.argv.slice(2)); }
    catch (error) { console.error('转换失败：' + error.message); process.exitCode = 1; }
}
module.exports = { cleanSpaces, prepareAst, formatTimestamp, publishOutput, openDocument, main };
