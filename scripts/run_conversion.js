
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

// ============ 配置：支持命令行参数 ============
// 从命令行获取源文件名，支持相对路径和绝对路径
const mdFileInput = process.argv[2];

if (!mdFileInput) {
    console.error('❌ 错误：请提供源 Markdown 文件名');
    console.error('用法: node run_conversion.js <源文件.md>');
    console.error('示例: node run_conversion.js 报告.md');
    process.exit(1);
}

// 解析源文件路径
const mdFile = path.resolve(mdFileInput);
if (!fs.existsSync(mdFile)) {
    console.error(`❌ 错误：文件不存在: ${mdFile}`);
    process.exit(1);
}

// 自动从 Skill 目录查找依赖文件
const SKILL_DIR = path.join(
    process.env.USERPROFILE || process.env.HOME,
    '.gemini', 'antigravity', 'skills', 'md2word-pandoc'
);

const referenceDoc = path.join(SKILL_DIR, 'templates', 'md2word模板.docx');
const filterScript = path.join(SKILL_DIR, 'scripts', 'style_filter.lua');

// 验证依赖文件
if (!fs.existsSync(referenceDoc)) {
    console.error(`❌ 错误：模板文件不存在: ${referenceDoc}`);
    console.error('   请确认 md2word-pandoc Skill 已正确安装');
    process.exit(1);
}

if (!fs.existsSync(filterScript)) {
    console.error(`❌ 错误：过滤器不存在: ${filterScript}`);
    process.exit(1);
}

// 生成输出文件名（基于源文件名 + 时间戳）
const now = new Date();
const offset = now.getTimezoneOffset() * 60000; // Beijing +8
const localDate = new Date(now.getTime() - offset);
const timestamp = localDate.toISOString().replace(/[:.]/g, '-').slice(0, 19);

// 获取源文件basename（不含扩展名）
const baseName = path.basename(mdFile, '.md');
const finalName = `${baseName}_Final_${timestamp}.docx`;

// 输出到源文件所在目录
const outputDir = path.dirname(mdFile);

// 临时文件（放在源文件目录，使用 ASCII 名避免路径问题）
const tmpInput = path.join(outputDir, 'temp_input.md');
const tmpOutput = path.join(outputDir, 'temp_output.docx');
const finalOutput = path.join(outputDir, finalName);

try {
    console.log("📄 源文件:", mdFile);
    console.log("📝 模板文件:", referenceDoc);
    console.log("🔧 过滤器:", filterScript);
    console.log("");

    console.log("1️⃣  准备文件 (自动清理格式)...");

    // 读取源文件内容
    let content = fs.readFileSync(mdFile, 'utf8');

    // 1. 清理 [汉字] [空格] [英文/数字]（仅同行内空白，不匹配换行符）
    // WHY: \s+ 会匹配 \n，导致标题末尾汉字与下一段首英文跨行合并
    content = content.replace(/([\u4e00-\u9fa5])[^\S\n\r]+([a-zA-Z0-9])/g, '$1$2');

    // 2. 清理 [英文/数字] [空格] [汉字]（仅同行内空白）
    content = content.replace(/([a-zA-Z0-9])[^\S\n\r]+([\u4e00-\u9fa5])/g, '$1$2');

    // 3. 尝试清理表格中的多余空行 (将连续两个换行符替换为一个，但在表格块内)
    // 注意：全篇替换可能会破坏段落结构，暂不激进处理，仅处理上述空格

    fs.writeFileSync(tmpInput, content, 'utf8');

    console.log("2️⃣  执行 Pandoc 转换...");
    // 使用引号包裹路径，防止空格导致的问题
    const cmd = `pandoc "${tmpInput}" -o "${tmpOutput}" --reference-doc="${referenceDoc}" --lua-filter="${filterScript}" --standalone`;
    console.log(`   执行命令: pandoc [源文件] -o [输出] --reference-doc=[模板] --lua-filter=[过滤器]`);
    execSync(cmd, { stdio: 'inherit' });

    // 2.5 合并封面+目录（模板同时作为样式源和封面内容源）
    const mergeScript = path.join(SKILL_DIR, 'scripts', 'merge_cover.py');
    if (fs.existsSync(mergeScript)) {
        console.log("2.5️⃣  合并封面与目录...");

        // 从 MD frontmatter 提取 title
        const titleMatch = content.match(/^---[\s\S]*?title:\s*(.+?)[\r\n]/m);
        const titleArg = titleMatch ? `--title "${titleMatch[1].trim()}"` : '';

        const mergeCmd = `python "${mergeScript}" "${referenceDoc}" "${tmpOutput}" "${tmpOutput}" ${titleArg}`;
        execSync(mergeCmd, { stdio: 'inherit' });
    }

    console.log("3️⃣  重命名输出文件...");
    if (fs.existsSync(tmpOutput)) {
        fs.renameSync(tmpOutput, finalOutput);
        console.log(`\n✅ 转换成功！\n`);
        console.log(`📦 输出文件: ${finalOutput}\n`);
    } else {
        throw new Error("Pandoc 未能生成输出文件");
    }

    // 清理临时文件
    fs.unlinkSync(tmpInput);

} catch (error) {
    console.error("\n❌ 转换失败:");
    console.error(error.message);

    // 清理临时文件
    if (fs.existsSync(tmpInput)) fs.unlinkSync(tmpInput);
    if (fs.existsSync(tmpOutput)) fs.unlinkSync(tmpOutput);

    process.exit(1);
}
