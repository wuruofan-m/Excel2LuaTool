local xlsPath = [[D:/SVN/public/cehua/db_excel]]
local luaPath = [[D:/SVN/lua]]

local M = {}

-- 分割字符串
local split = function(str, split_char)
    -- 以下3行代码做数据校检（在底端设备上尽量去掉）
    if type(str) ~= "string" or #str == 0 then
        return {}
    end
    local nSepLen = string.len(split_char)
    local sub_str_tab = {}
    while (true) do
        local pos = string.find(str, split_char)
        if (not pos) then
            sub_str_tab[#sub_str_tab + 1] = str
            break
        end
        local sub_str = string.sub(str, 1, pos - 1)
        sub_str_tab[#sub_str_tab + 1] = sub_str
        str = string.sub(str, pos + nSepLen, #str)
    end
    return sub_str_tab
end

-- 从后面开始查找
local findStrFromBehind = function(str, findstr)
    local len = string.len(str)
    local val = string.len(findstr) - 1
    for i = len, 1, -1 do
        if i + val <= len then
            local temp = string.sub(str, i, i + val)
            if temp == findstr then
                return i, i + val
            end
        end
    end
    return "nil"
end

M.parsing_excel = function(filePath)
    require('luacom')
    local excel = luacom.CreateObject("Excel.Application") -- 创建新的实例
    if excel == nil then
        error("Object is not create")
    end

    -- excel.Visible = false
    -- 打开指定的文件
    local book = excel.Workbooks:Open(filePath)
    if not book then
        error("book is nil")
    end
    -- 读第一个工作表
    local sheet = book.Sheets(1)

    local keys = {}
    local cols = {}
    local type_yn = {} -- { {"string","yes"}...} string类型，true = 必须填写数值

    -- 获取相应的 key值和 列号
    for i = 1, 30 do
        local key = sheet.Cells(3, i).Value2
        if key then
            table.insert(keys, key)
            table.insert(cols, i)
            local ty = sheet.Cells(2, i).Value2
            local slt = string.find(ty, "[%p]")
            table.insert(type_yn, split(ty, string.sub(ty, slt, slt)))
        end
    end

    -- 获取数据内容 并检测
    local content = {}
    local row = 4
    local flag = true
    while flag do
        local tempRow = {}
        for k, col in pairs(cols) do
            if not sheet.Cells(row, cols[1]).Value2 then
                flag = false
                break
            end
            local value = sheet.Cells(row, col).Value2
            if M.checkValue(type_yn[k], value) then
                -- 中文问题

                value = value == "nil" and "nil" or value
                table.insert(tempRow, value)
                -- print("row:"..row,"col:"..col,"value:"..tostring(value))
            else
                print(string.format("%s : %d 行 %d列", filePath, row, col))
                error(string.format("%s : row = %d , col = %d列", filePath, row, col))
            end
        end
        if tempRow and #tempRow ~= 0 then
            table.insert(content, tempRow)
        end
        row = row + 1
    end
    row = 4

    book:Close()
    excel:Quit()
    excel = nil
    collectgarbage()

    return {
        keys = keys,
        type_yn = type_yn,
        content = content
    }
end

-- 查找 xls 文件
require "lfs"
M.findInDir = function(path, wefind, r_table, intofolder)

    for file in lfs.dir(path) do
        if file ~= "." and file ~= ".." then
            local f = path .. '/' .. file

            if string.find(f, wefind) ~= nil then

                table.insert(r_table, f)
            end
            local attr = lfs.attributes(f)
            assert(type(attr) == "table")
            if attr.mode == "directory" and intofolder then
                M.findInDir(f, wefind, r_table, intofolder)
            else
                -- for name, value in pairs(attr) do
                --    print (name, value)
                -- end
            end
        end
    end
end

-- 存为lua 文件
M.saveLuaFile = function(path, data)

    path = string.gsub(path, xlsPath, luaPath)
    local b, e = findStrFromBehind(path, "/")
    local dotp = findStrFromBehind(path, ".")
    local fileName = "DB_" .. string.sub(path, b + 1, dotp - 1)
    path = string.sub(path, 1, b) .. fileName .. ".lua"

    local file = io.open(path, 'w+')
    if not file then
        local dirPath = string.sub(path, 1, b)
        dirPath = string.gsub(dirPath, "/", "\\\\")
        os.execute("md " .. dirPath)
        file = io.open(path, 'w+')
    end

    M.setLuaContent(file, data, fileName)
    file:close()
end

-- lua 文件内容编写
M.setLuaContent = function(file, data, fileName)
    local keys, content, type_yn = data.keys, data.content, data.type_yn
    local tab = "	"
    local str = ""
    -- 注释内容
    str = str .. "-- Filename: " .. tostring(fileName) .. ".lua\n"
    str = str .. "-- Author: auto-created by kyn`s ParseExcel(to lua) tool." .. "\n"
    str = str .. "-- methods: X.keys = {}, X.getDataById(id), X.getArrDataByField(fieldName, fieldValue)" .. "\n"
    str = str .. "-- Function: no description." .. "\n\n"

    -- key
    str = str .. "local keys = {\n" .. tab

    for k, v in pairs(keys) do
        str = str .. "\"" .. tostring(v) .. "\"" .. ","
    end
    str = str .. "\n}\n\n"

    -- data
    str = str .. "local data = {\n"

    for row, v in pairs(content) do
        str = str .. tab .. "id" .. row .. " = {"
        for col, val in pairs(v) do
            str = str .. M.buildString(type_yn[col], val) .. ","
        end

        str = str .. "}\n"
    end

    str = str .. "}\n\n"

    -- return

    str = str .. tostring(fileName) .. " = DB_Template.new (keys, data)\n"
    str = str .. "return " .. tostring(fileName)

    -- print(str)
    file:write(str)
end

-- 检测 类型，数据是否存在
M.checkValue = function(type_yn, value)
    -- 判断是否存在
    if type_yn[2] == "yes" then

        if value == nil or value == "nil" then
            return false
        end
    else
        if value == nil or value == "nil" then
            return true
        end
    end

    if type_yn[1] == "string" or type_yn[1] == "String" then

    end

    if type_yn[1] == "number" or type_yn[1] == "Number" then
        local b = string.find(value, "[^0-9]")
        if b and string.sub(value, b, b) ~= "." and string.sub(value, b, b) ~= "-" then
            return false
        end
    end

    if string.find(type_yn[1], "array") and value ~= "nil" then
        local t = split(value, ",")
        if #t <= 1 then
            return false
        end
    end

    return true
end

-- 构造 字符串，用于写于lua脚本
M.buildString = function(type_yn, value)

    local str = ""

    if value == "nil" then
        str = value
        return str
    end

    if type_yn[1] == "Number" or type_yn[1] == "number" then
        str = value
        return str
    end

    if type_yn[1] == "string" or type_yn[1] == "String" then
        str = "\"" .. value .. "\""
        return str
    end

    if string.find(type_yn[1], "array") then
        local t = split(tostring(value), ",")
        local str = "{"
        for k, v in pairs(t) do
            str = str .. v
            str = str .. (k ~= #t and "," or "")
        end
        str = str .. "}"
        return str
    end

end

------------------------------------------------------------------------------------------
local xlsList = {}
M.findInDir(xlsPath, "%.xls", xlsList, true)

local lastClock = os.clock()
for k, v in pairs(xlsList) do
    local data = M.parsing_excel(v)
    M.saveLuaFile(v, data)

    local curClock = os.clock()
    print(k, v, curClock - lastClock)
    lastClock = curClock

end

print("total time:" .. os.clock())

-- 每个文件大概在 1-2秒之间，根据数据量，数据越多，时间越长
-- 优化方案：对每个文件内容进行MD5，存起来，下次做对比，内容没有改变的，直接跳过。
