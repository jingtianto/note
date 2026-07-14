; ============================================
;  双击此 AHK 脚本 → 自动注入 CSS 样式到浏览器控制台
;  支持 Chrome / Edge（其他浏览器提示手动粘贴）
; ============================================

; 准备要粘贴的 JavaScript 代码
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

; 复制到剪贴板
clipboard = %jsCode%

; 获取当前活动窗口的进程名
WinGet, pid, PID, A
WinGet, exe, ProcessName, ahk_pid %pid%

; 判断是否为 Chrome 或 Edge
if (exe = "chrome.exe" or exe = "msedge.exe") {
    ; 发送 Ctrl+Shift+J 打开控制台 (Chrome/Edge)
    Send, ^+j
    Sleep, 300
    Send, ^v
    Sleep, 200
    Send, {Enter}
    MsgBox, 代码已粘贴并执行。请查看控制台输出。
} else {
    MsgBox, 当前窗口不是 Chrome 或 Edge，代码已复制到剪贴板。`n请手动打开控制台 (Ctrl+Shift+J) 后粘贴执行。
}