#Requires AutoHotkey >=2.0

; ---------- 准备 JavaScript 代码 ----------
jsCode := "
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
)"

; ---------- 复制到剪贴板 ----------
A_Clipboard := jsCode

; ---------- 显示倒计时提示 ----------
ToolTip("请在 5 秒内切换到目标 Edge/Chrome 窗口...", 0, 0)
Sleep(1000)
ToolTip("剩余 4 秒...", 0, 0)
Sleep(1000)
ToolTip("剩余 3 秒...", 0, 0)
Sleep(1000)
ToolTip("剩余 2 秒...", 0, 0)
Sleep(1000)
ToolTip("剩余 1 秒...", 0, 0)
Sleep(1000)
ToolTip()  ; 清除提示

; ---------- 获取当前活动窗口的进程名 ----------
currentProcess := WinGetProcessName("A")

; ---------- 检查是否为 Edge 或 Chrome ----------
if (currentProcess = "msedge.exe" or currentProcess = "chrome.exe") {
    ; 当前窗口是浏览器，执行注入
    Send("^+j")          ; 打开控制台 (Chrome/Edge)
    Sleep(300)
    Send("^v")           ; 粘贴
    Sleep(200)
    Send("{Enter}")      ; 执行
    MsgBox("✅ 样式已注入！请查看控制台输出。", "成功", "OK Iconi")
} else {
    MsgBox("当前活动窗口不是 Edge 或 Chrome。`n请确保在 5 秒内切换到了目标浏览器窗口。", "提示", "OK Iconi")
}