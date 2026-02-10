-- style_filter.lua
-- 自动检测 H1 数量，决定标题偏移策略：
--   H1 = 1: 视为文档总标题，转为 Title 样式，H2→Heading1
--   H1 >= 2: 不偏移，H1→Heading1

-- 全局变量：存储 H1 统计和偏移策略
local h1_count = 0
local shift_mode = nil  -- "shift" or "no_shift"

-- Helper function to recursively find the first Str element and its parent list
local function get_first_str_ref(content)
  if #content == 0 then return nil, nil, nil end
  for i, item in ipairs(content) do
    if item.t == 'Str' then
      return item, content, i
    elseif item.t == 'Strong' or item.t == 'Emph' or item.t == 'Span' then
      local found_item, found_list, found_index = get_first_str_ref(item.content)
      if found_item then
        return found_item, found_list, found_index
      end
    end
    if i == 1 then
       if item.t == 'Space' then
         -- continue searching next sibling
       else
         return nil, nil, nil
       end
    end
  end
  return nil, nil, nil
end

-- 处理单个标题：去除手动编号
local function strip_manual_numbering(el)
  local str_item, parent_list, index = get_first_str_ref(el.content)
  
  if str_item then
    local text = str_item.text
    local processed = false
    
    -- Case 1: Chinese Chapter "第1章" or "第一章"
    -- Match "第" followed by non-whitespace chars (numbers or hanzi) then "章"
    if text:match("^第%S*章") then
       text = text:gsub("^第%S*章%s*", "")
       processed = true
       
    -- Case 2: Numeric with dot "1." or "1.1" or "1.1.1"
    elseif text:match("^%d+%.") then
       text = text:gsub("^%d+%.[%d%.]*%s*", "")
       processed = true

    -- Case 3: Numeric with Dunhao "1、"
    elseif text:match("^%d+、") then
       text = text:gsub("^%d+、%s*", "")
       processed = true

    -- Case 4: Parentheses "(1)"
    elseif text:match("^%(%d+%)") then
       text = text:gsub("^%(%d+%)%s*", "")
       processed = true

    -- Case 5: Numeric with Space "1 " (but allow "2026 " year context)
    -- Strict check: Start with digit(s), followed by space, but ensure it's not a year (4 digits)
    elseif text:match("^%d+%s+") then
       -- Simple heuristic: if it's 4 digits, assume it might be a year, don't strip
       -- if it's 1-3 digits, strip it.
       if not text:match("^%d%d%d%d%s+") then
           text = text:gsub("^%d+%s+", "")
           processed = true
       end
    end
    
    if processed then
       str_item.text = text
       if text == "" then
         table.remove(parent_list, index)
         if #parent_list >= index and parent_list[index].t == 'Space' then
             table.remove(parent_list, index)
         end
       end
    end
  end
end

-- 主过滤器：处理整个文档
function Pandoc(doc)
  -- Step 1: 统计 H1 数量
  for _, block in ipairs(doc.blocks) do
    if block.t == 'Header' and block.level == 1 then
      h1_count = h1_count + 1
    end
  end
  
  -- Step 2: 决定偏移策略
  -- H1 = 0: 文档从 H2 开始，需要偏移 (H2->Heading1)
  -- H1 = 1: 唯一 H1 作为文档标题，需要偏移 (H1->Title, H2->Heading1)
  -- H1 >= 2: H1 是正式章节，不偏移 (H1->Heading1)
  if h1_count <= 1 then
    shift_mode = "shift"
    print("[style_filter] H1 count = " .. h1_count .. " -> Shift mode (H2->Heading1)")
  else
    shift_mode = "no_shift"
    print("[style_filter] H1 count = " .. h1_count .. " -> No shift mode (H1->Heading1)")
  end
  
  -- Step 3: 处理所有标题
  local new_blocks = {}
  for _, block in ipairs(doc.blocks) do
    if block.t == 'Header' then
      if shift_mode == "shift" then
        -- H1 = 1 的情况：H1 转为 Title 样式，其他降级
        if block.level == 1 then
          -- 将 H1 转为带 Title 样式的 Div
          local title_div = pandoc.Div(
            {pandoc.Para(block.content)},
            {['custom-style'] = 'Title'}
          )
          table.insert(new_blocks, title_div)
        else
          -- H2->H1, H3->H2, etc.
          block.level = block.level - 1
          strip_manual_numbering(block)
          table.insert(new_blocks, block)
        end
      else
        -- H1 >= 2 的情况：不偏移
        strip_manual_numbering(block)
        table.insert(new_blocks, block)
      end
    else
      table.insert(new_blocks, block)
    end
  end
  
  doc.blocks = new_blocks

  -- 表格内容样式：将 Plain/Para 包装为带 custom-style 的 Div
  -- WHY: Table 函数不能与 Pandoc 函数共存，必须在此处通过 walk 处理
  doc = doc:walk {
    Table = function(tbl)
      return tbl:walk {
        Plain = function(p)
          return pandoc.Div({pandoc.Para(p.content)}, {['custom-style'] = 'TableContent'})
        end
      }
    end
  }

  return doc
end
