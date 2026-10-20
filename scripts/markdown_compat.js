// Compatibility with the existing Markdown writing and cleanup conventions.
function textOf(value) {
    if (Array.isArray(value)) return value.map(textOf).join('');
    if (!value || typeof value !== 'object') return '';
    if (value.t === 'Str') return value.c;
    if (['Space', 'SoftBreak', 'LineBreak'].includes(value.t)) return ' ';
    return textOf(value.c);
}

function restoreImageCaptions(ast) {
    const output = [];
    for (let i = 0; i < ast.blocks.length; i++) {
        const block = ast.blocks[i];
        output.push(block);
        if (!['Para', 'Plain'].includes(block.t)) continue;
        const inlines = block.c.filter(item => !['Space', 'SoftBreak', 'LineBreak'].includes(item.t));
        if (inlines.length !== 1 || inlines[0].t !== 'Image') continue;
        const alt = inlines[0].c[1];
        const title = textOf(alt).trim();
        if (!title) continue;
        const next = ast.blocks[i + 1];
        const explicit = next && ['Para', 'Plain'].includes(next.t) ? textOf(next.c).trim() : '';
        // An explicit caption takes precedence over the image's alternative text.
        if (/^图\s*(?:\d|[、．.:：\-–—\s])/.test(explicit) || explicit === title) continue;
        output.push({ t: 'Para', c: JSON.parse(JSON.stringify(alt)) });
    }
    ast.blocks = output;
    return ast;
}

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
    content = content.replace(/(```[\s\S]*?```|~~~[\s\S]*?~~~)/g, (match) => {
        protectedBlocks.push(match);
        return `\x00PROT_${protectedBlocks.length - 1}\x00`;
    });
    content = content.replace(/(\$\$[\s\S]*?\$\$)/g, (match) => {
        protectedBlocks.push(match);
        return `\x00PROT_${protectedBlocks.length - 1}\x00`;
    });

    // Keep inline code, math and Markdown resource references unchanged.
    for (const pattern of [ /(`+)[^\n]*?\1/g, /\$(?!\$)(?:\\.|[^$\n])+\$/g,
        /\\\[[\s\S]*?\\\]/g, /!?\[[^\]\n]*\]\([^\n]*?\)/g ]) {
        content = content.replace(pattern, match => {
            protectedBlocks.push(match);
            return `\x00PROT_${protectedBlocks.length - 1}\x00`;
        });
    }

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

module.exports = { cleanSpaces, restoreImageCaptions };
