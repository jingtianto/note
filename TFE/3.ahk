; ============================================================
;  双击此 AHK 脚本 → 自动将 CSS 样式 JS 代码粘贴到浏览器控制台
;  支持 Chrome / Edge / Chromium 内核浏览器
;  要求：浏览器已打开且为当前活动窗口
; ============================================================

#NoEnv
#SingleInstance, Force
SetWorkingDir %A_ScriptDir%

; ---------- 准备要粘贴的 JavaScript 代码 ----------
jsCode =
(
(function(){
    const style = document.createElement('style');
    style.textContent = `
.ant-tooltip-inner {
    background-color: #202630 !important;
    opacity: 1 !important;
    font-family: "Segoe UI", Roboto, sans-serif !important;
    color: #ffffff !important;
    font-weight: 400 !important;
    font-size: 14px !important;
    line-height: 1.55 !important;
    padding: 10px 14px !important;
    border-radius: 4px !important;
    -webkit-font-smoothing: antialiased !important;
    -moz-osx-font-smoothing: grayscale !important;
    text-shadow: none !important;
}
.ant-tooltip-inner *,
.ant-tooltip-inner .ant-card,
.ant-tooltip-inner .ant-card-head-title,
.ant-tooltip-inner .ant-card-body span {
    font-family: "Segoe UI", Roboto, sans-serif !important;
    color: #ffffff !important;
    font-weight: 400 !important;
    -webkit-font-smoothing: antialiased !important;
    -moz-osx-font-smoothing: grayscale !important;
    text-shadow: none !important;
}
.ant-tooltip-placement-top .ant-tooltip-arrow::before {border-top-color: #202630 !important;}
.ant-tooltip-placement-bottom .ant-tooltip-arrow::before {border-bottom-color: #202630 !important;}
.ant-tooltip-placement-left .ant-tooltip-arrow::before {border-left-color: #202630 !important;}
.ant-tooltip-placement-right .ant-tooltip-arrow::before {border-right-color: #202630 !important;}
`;
    document.head.appendChild(style);
    console.log("✅ 已生效：常规字重不加粗，字体锐化消除重影");
})();
)

; ---------- 复制到剪贴板 ----------
Clipboard := jsCode

; ---------- 检查当前窗口是否为浏览器 ----------
WinGetClass, winClass, A
if (winClass ~= "Chrome_WidgetWin_1|Edge|Chrome|Chromium") {
    ; 发送 Ctrl+Shift+J 打开 Chrome/Edge 的 Console
    Send, ^+j
    Sleep, 200
    ; 粘贴并执行
    Send, ^v
    Sleep, 100
    Send, {Enter}
    ToolTip, ✅ 代码已粘贴并执行
} else {
    ; 非浏览器窗口，仅复制到剪贴板并提示
    ToolTip, ℹ️ 代码已复制到剪贴板，请手动粘贴到浏览器控制台（Ctrl+Shift+J）
}
Sleep, 1500
ToolTip  ; 移除提示
return