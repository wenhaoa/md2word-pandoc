
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

// ============ 配置：支持命令行参数 ============
// 解析命令行参数
const args = process.argv.slice(2);
const openAfterConvert = args.includes('--open');
const skipCaption = args.includes('--no-caption');
const mdFileInput = args.find(a => !a.startsWith('--'));

if (!mdFileInput) {
    console.error('❌ 错误：请提供源 Markdown 文件名');
    console.error('用法: node run_conversion.js <源文件.md> [--open] [--no-caption]');
    console.error('示例: node run_conversion.js 报告.md --open');
    console.error('  --no-caption    跳过图表题注 SEQ 域处理');
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
const finalName = `${baseName}_${timestamp}.docx`;

// 输出到源文件所在目录
const outputDir = path.dirname(mdFile);

// 临时文件（放在源文件目录，使用 ASCII 名避免路径问题）
const tmpInput = path.join(outputDir, 'temp_input.md');
const tmpOutput = path.join(outputDir, 'temp_output.docx');
const finalOutput = path.join(outputDir, finalName);

// ============ 空格清理：对一行正文执行多条规则 ============
function cleanLineBody(body) {
    // 规则1: [CJK字符/标点] + 正好 1 个空格 + [任意非空白] → 删除空格
    body = body.replace(/([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])[ \t](\S)/g, '$1$2');

    // 规则2: [任意非空白] + 正好 1 个空格 + [CJK字符/标点] → 删除空格
    body = body.replace(/(\S)[ \t]([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])/g, '$1$2');

    // 规则3: [数字] + 正好 1 个空格 + [单位字母/符号] → 删除空格 (1rad, 5mm)
    body = body.replace(/(\d)[ \t]([a-zA-Z\u00b0\u00b5\u03bc%\u2030])/g, '$1$2');

    // 规则4: [比较符] + 正好 1 个空格 + [数字] → 删除空格 (<1, >5)
    body = body.replace(/([\u003c\u003e\u2264\u2265\u2248])[ \t](\d)/g, '$1$2');

    // 规则5: 逗号 + 正好 1 个空格 + [数字/正负号] → 删除空格 (坐标紧凑化)
    //   覆盖：(0.00, 1346.222) → (0.00,1346.222)
    body = body.replace(/,[ \t]([+\-\u2212]?\d)/g, ',$1');

    return body;
}

function cleanSpaces(content) {
    // Step 1: 保护代码围栏、显示公式和 YAML frontmatter
    const protectedBlocks = [];
    // 保护 YAML frontmatter（文件开头的 ---...--- 块）
    content = content.replace(/^---[\s\S]*?\n---/m, (match) => {
        protectedBlocks.push(match);
        return `\x00PROT_${protectedBlocks.length - 1}\x00`;
    });
    content = content.replace(/(```[\s\S]*?```)/g, (match) => {
        protectedBlocks.push(match);
        return `\x00PROT_${protectedBlocks.length - 1}\x00`;
    });
    content = content.replace(/(\$\$[\s\S]*?\$\$)/g, (match) => {
        protectedBlocks.push(match);
        return `\x00PROT_${protectedBlocks.length - 1}\x00`;
    });

    // Step 1.3: 中文双引号转换（在保护区生效后、空格清理前执行）
    // WHY: 成对匹配 "..." 并要求内容包含至少一个 CJK 字符，
    // 这样 YAML 中的 "value"、英文引号都不受影响。
    // 在空格清理前执行，避免空格删除后丢失引号方向信息。
    content = content.replace(/"([^"\n]*?[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef][^"\n]*?)"/g, '\u201C$1\u201D');

    // Step 1.5: 连接 CJK 跨行软换行
    // WHY: 当一行以 CJK 字符/标点结尾，下一行以 CJK 字符/标点开头时，
    // 删除中间的换行符，防止 Pandoc 在此处插入空格
    content = content.replace(
        /([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])\r?\n([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])/g,
        '$1$2'
    );

    // Step 2: 逐行处理
    const lines = content.split('\n');
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];

        // 跳过含占位符的行
        if (line.includes('\x00PROT_')) continue;

        // 跳过表格分隔行 |---|---|
        if (/^\|[\s\-:|]+\|/.test(line)) continue;

        // --- 提取行首 Markdown 语法前缀 (不参与清理) ---
        let prefix = '';
        let body = line;

        // 标题行：保护整个 "# " / "## " / "### 1.2.3 " / "## 第X章 " 前缀
        const headerMatch = line.match(/^(#{1,6}\s+(?:(?:\d+\.)+\d*\s+)?(?:\u7b2c[\S]*\u7ae0\s+)?)/);
        if (headerMatch) {
            prefix = headerMatch[1];
            body = line.slice(prefix.length);
        } else {
            // 引用块 > 
            const quoteMatch = line.match(/^(>\s*)/);
            if (quoteMatch) {
                prefix = quoteMatch[1];
                body = line.slice(prefix.length);
            } else {
                // 列表项前缀 (如 "- ", "* ", "1. ")
                const listMatch = line.match(/^([ \t]*[\*\-\+][ \t]+|[ \t]*\d+\.[ \t]+)/);
                if (listMatch) {
                    prefix = listMatch[1];
                    body = line.slice(prefix.length);
                }
            }
        }

        // --- 防御性判定：如果 body 中存在 `\S` 之间至少 2 个空格的对齐块 ---
        // WHY: 极大概率是 Pandoc Simple/Multiline/Grid Table 的列分隔符，
        // 跳过整行避免破坏对齐
        if (/\S[ \t]{2,}\S/.test(body)) {
            continue;
        }

        // 对 body 执行清理
        body = cleanLineBody(body);

        lines[i] = prefix + body;
    }
    content = lines.join('\n');

    // Step 3: 恢复受保护的块
    content = content.replace(/\x00PROT_(\d+)\x00/g, (_, idx) => protectedBlocks[parseInt(idx)]);

    return content;
}

try {
    console.log("📄 源文件:", mdFile);
    console.log("📝 模板文件:", referenceDoc);
    console.log("🔧 过滤器:", filterScript);
    console.log("");

    console.log("1️⃣  准备文件 (自动清理格式)...");

    // 读取源文件内容
    let content = fs.readFileSync(mdFile, 'utf8');
    content = cleanSpaces(content);


    fs.writeFileSync(tmpInput, content, 'utf8');

    console.log("2️⃣  执行 Pandoc 转换...");
    // 使用引号包裹路径，防止空格导致的问题
    // WHY: --resource-path 确保右键发送时也能找到相对路径图片
    // WHY: -subscript-superscript 防止 ~ 和 ^ 被解析为下标/上标（中文常用 300~400km 表示范围）
    const cmd = `pandoc "${tmpInput}" -o "${tmpOutput}" --from markdown-smart-subscript-superscript --reference-doc="${referenceDoc}" --lua-filter="${filterScript}" --resource-path="${outputDir}" --standalone`;
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

    // 2.6 添加图表题注 SEQ 域
    const captionScript = path.join(SKILL_DIR, 'scripts', 'add_captions.py');
    if (!skipCaption && fs.existsSync(captionScript)) {
        console.log("2.6️⃣  添加图表题注 SEQ 域...");
        try {
            execSync(`python "${captionScript}" "${tmpOutput}"`, { stdio: 'inherit' });
        } catch (e) {
            console.warn('  ⚠ 题注处理失败（非致命）:', e.message);
        }
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

    // WHY: --open 标志用于 GUI 入口（BAT 双击/拖拽），命令行模式默认不打开
    if (openAfterConvert) {
        console.log('📖 正在打开 Word 文件...');
        const { exec } = require('child_process');
        exec(`start "" "${finalOutput}"`);
    }

} catch (error) {
    console.error("\n❌ 转换失败:");
    console.error(error.message);

    // 清理临时文件
    if (fs.existsSync(tmpInput)) fs.unlinkSync(tmpInput);
    if (fs.existsSync(tmpOutput)) fs.unlinkSync(tmpOutput);

    process.exit(1);
}
