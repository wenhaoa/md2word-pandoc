const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { formatTimestamp, publishOutput, openDocument } = require('./run_conversion.js');

assert.equal(formatTimestamp(Date.parse('2026-05-07T06:48:22.000Z')), '20260507_144822');

// Verify the observable output contract using the real converter.
const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'md2word-filename-'));
try {
    const source = path.join(scratch, '示例技术报告.md');
    const content = '---\ntitle: 示例技术报告\n---\n\n## 1. 概述\n\n正文。\n';
    fs.writeFileSync(source, content, 'utf8');
    const before = Date.now();
    execFileSync(process.execPath, [path.join(__dirname, 'run_conversion.js'), source],
        { env: { ...process.env, TZ: 'UTC', PYTHONIOENCODING: 'utf-8', PYTHONDONTWRITEBYTECODE: '1' }, encoding: 'utf8' });
    const after = Date.now();
    const outputs = fs.readdirSync(scratch).filter(name => name.endsWith('.docx'));
    assert.equal(outputs.length, 1);
    const match = outputs[0].match(/^示例技术报告(\d{8}_\d{6})\.docx$/);
    assert.ok(match, 'filename must keep the original compact timestamp without an extra suffix');
    const stamp = formatTimestamp;
    assert.ok(match[1] >= stamp(before) && match[1] <= stamp(after), 'timestamp must use Beijing time even with TZ=UTC');
    assert.equal(fs.readFileSync(source, 'utf8'), content, 'source must remain intact');
    assert.equal(fs.readdirSync(scratch).filter(name => name.startsWith('.md2word-')).length, 0);
    const first = publishOutput(path.join(scratch, outputs[0]), scratch, '并发报告', '20260507_144822');
    const second = publishOutput(path.join(scratch, outputs[0]), scratch, '并发报告', '20260507_144822');
    assert.equal(path.basename(first), '并发报告20260507_144822.docx');
    assert.equal(path.basename(second), '并发报告20260507_144822_2.docx');
    assert.deepEqual(fs.readFileSync(first), fs.readFileSync(second));
    if (process.platform === 'win32') {
        const unusualPath = path.join(scratch, "中文 & '报告' [1].docx");
        let called = false;
        openDocument(unusualPath, process.env, (program, args, options) => {
            called = true;
            assert.equal(program, 'powershell.exe');
            assert.equal(options.env.MD2WORD_OPEN_FILE, unusualPath);
            assert.ok(!args.at(-1).includes(unusualPath));
        });
        assert.ok(called, 'the launcher must finish before openDocument returns');
        assert.throws(() => openDocument(unusualPath, process.env, () => {
            throw new Error('association launch failed');
        }), /association launch failed/);
    }
    console.log('filename, collision and open-launch tests passed');
} finally {
    if (path.dirname(scratch) !== os.tmpdir() || !path.basename(scratch).startsWith('md2word-filename-')) {
        throw new Error('unexpected temporary directory; preserving it');
    }
    fs.rmSync(scratch, { recursive: true, force: true });
}
