#Requires AutoHotkey v2.0

; 获取当天日期，格式为 DD.MM.YYYY
today := FormatTime(, "dd.MM.yyyy")

; 拼接目标文本
text_to_copy := today ",qxz715a: passed. Automatically passed."

; 写入剪贴板
A_Clipboard := text_to_copy

; 可选：弹出提示，告诉你复制成功了
MsgBox("Copied：`n" text_to_copy, "AutoHotkey", 0x40)
