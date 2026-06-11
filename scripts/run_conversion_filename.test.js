const assert = require('assert');
const path = require('path');
const vm = require('vm');
const fs = require('fs');

const source = fs.readFileSync(path.join(__dirname, 'run_conversion.js'), 'utf8');
const renamed = [];
const written = [];
const commands = [];

class FixedDate extends Date {
    constructor(...args) {
        if (args.length === 0) {
            super('2026-05-07T06:48:22.000Z');
        } else {
            super(...args);
        }
    }
}

const fakeFs = {
    existsSync(filePath) {
        return [
            path.resolve('示例技术报告.md'),
            path.resolve(__dirname, '..', 'templates', 'md2word模板.docx'),
            path.resolve(__dirname, 'style_filter.lua'),
            path.resolve(__dirname, 'merge_cover.py'),
            path.resolve(__dirname, 'add_captions.py'),
            path.resolve('temp_output.docx'),
        ].includes(path.resolve(filePath));
    },
    readFileSync() {
        return '---\ntitle: 示例技术报告\n---\n\n## 1. 概述\n';
    },
    writeFileSync(filePath, content) {
        written.push({ filePath, content });
    },
    renameSync(from, to) {
        renamed.push({ from, to });
    },
    unlinkSync() {},
};

const sandbox = {
    __dirname,
    console: { log() {}, error() {}, warn() {} },
    process: {
        argv: ['node', 'run_conversion.js', '示例技术报告.md'],
        env: { USERPROFILE: path.resolve(__dirname, '..') },
        exit(code) { throw new Error(`process.exit(${code})`); },
    },
    Date: FixedDate,
    require(moduleName) {
        if (moduleName === 'fs') return fakeFs;
        if (moduleName === 'child_process') return {
            execSync(command) { commands.push(command); },
            exec() {},
        };
        return require(moduleName);
    },
};

vm.runInNewContext(source, sandbox, { filename: 'run_conversion.js' });

assert.strictEqual(renamed.length, 1, 'conversion should rename the temporary output once');
assert.strictEqual(
    path.basename(renamed[0].to),
    '示例技术报告20260507_144822.docx',
    'output filename should use compact local timestamp with an underscore between date and time'
);
assert.strictEqual(written.length, 1, 'conversion should still write the preprocessed markdown once');
assert(
    commands.some(command => command.includes('--date-cn "2026 年 5 月"')),
    'merge command should pass the local Chinese year-month cover date'
);

console.log('run_conversion filename tests passed');
