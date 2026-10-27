// Explicit, single-table page layout; never infer orientation from column count.
function countTables(value) {
    if (Array.isArray(value)) return value.reduce((sum, item) => sum + countTables(item), 0);
    if (!value || typeof value !== 'object') return 0;
    return (value.t === 'Table' ? 1 : 0) + Object.values(value).reduce((sum, item) => sum + countTables(item), 0);
}

function isLandscape(value) {
    return value && value.t === 'Div' && value.c[0][1].includes('landscape');
}

function hasNestedMarker(value) {
    if (Array.isArray(value)) return value.some(hasNestedMarker);
    if (!value || typeof value !== 'object') return false;
    return isLandscape(value) || Object.values(value).some(hasNestedMarker);
}

function prepareTableLayouts(ast) {
    let tableCount = 0;
    const landscape = [];
    ast.blocks = ast.blocks.map(block => {
        if (isLandscape(block)) {
            const contents = block.c[1];
            if (contents.length !== 1 || contents[0].t !== 'Table' || hasNestedMarker(contents[0])) {
                throw new Error('landscape块只能包含一张表格，且不能嵌套横向标记');
            }
            landscape.push(tableCount);
            tableCount += countTables(contents[0]);
            return contents[0];
        }
        if (hasNestedMarker(block)) throw new Error('landscape块必须位于正文顶层');
        tableCount += countTables(block);
        return block;
    });
    return { table_count: tableCount, landscape };
}

module.exports = { prepareTableLayouts };
