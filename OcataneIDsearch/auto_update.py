#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
独立增量更新脚本 - 用于定时任务
功能：读取 config.json，对所有启用的数据源执行增量更新
运行方式：python auto_update.py
支持显示详细进度，包括目录扫描进度
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

# 导入主程序中的 IndexBuilder 类（不会触发 GUI 界面）
try:
    from OcatanID import IndexBuilder
except ImportError:
    print("错误：找不到 OcatanID.py，请确保该文件在同一目录下")
    sys.exit(1)

# ===== 进度回调函数 =====
def progress_callback(msg_type, *args):
    """
    回调函数，用于接收索引构建过程中的状态信息并打印到控制台
    """
    if msg_type == "status":
        # 状态信息，例如“正在扫描: /path”
        print(f"   📌 {args[0]}")
    
    elif msg_type == "progress":
        # 进度信息：当前索引、总数、文件名
        idx, total, name = args
        pct = int((idx / total) * 100) if total > 0 else 0
        print(f"   ⏳ [{idx}/{total}] ({pct}%) {name}")
    
    elif msg_type == "debug":
        # 调试信息：包括目录扫描进度（每100个目录输出一次）
        print(f"   🔍 {args[0]}")
    
    elif msg_type == "done":
        # 完成信息
        print(f"   ✅ {args[0]}")
    
    elif msg_type == "error":
        # 错误信息
        print(f"   ❌ {args[0]}")
    
    elif msg_type == "source":
        # 数据源切换
        print(f"\n📁 {args[0]}")
    
    elif msg_type == "result":
        # 结果信息（由主循环单独打印，此处忽略避免重复）
        pass

# ===== 主函数 =====
def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")
    
    if not os.path.exists(config_path):
        print(f"错误：未找到 config.json（路径：{config_path}）")
        sys.exit(1)
    
    # 读取配置，获取启用的数据源
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    sources = config.get("data_sources", [])
    enabled_sources = [s for s in sources if s.get("enabled", True)]
    
    if not enabled_sources:
        print("没有启用的数据源，无需更新")
        return
    
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 开始增量更新...")
    print(f"共 {len(enabled_sources)} 个数据源需要更新")
    
    # 初始化构建器（传入 config_path，它会自动加载配置）
    builder = IndexBuilder(config_path)
    
    success_count = 0
    for source in enabled_sources:
        name = source["name"]
        print(f"\n▶ 正在更新数据源: {name}")
        
        # 执行增量更新（incremental=True），传入回调
        result = builder.build_index(
            source_name=name,
            output_dir=script_dir,          # JSON 文件输出到当前目录
            callback=progress_callback,     # 显示进度
            incremental=True
        )
        
        # 打印结果摘要
        if "error" in result:
            print(f"   ❌ 更新失败: {result['error']}")
        elif result.get("stopped"):
            print(f"   ⏹ 已停止")
        else:
            print(f"   ✅ 更新完成: 新增 {result.get('new', 0)} 个文件, "
                  f"修改 {result.get('modified', 0)} 个文件, "
                  f"删除 {result.get('deleted', 0)} 个文件, "
                  f"当前索引含 {result.get('total_ids', 0)} 个唯一 ID")
            success_count += 1
    
    print(f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 全部完成！成功更新 {success_count}/{len(enabled_sources)} 个数据源")

if __name__ == "__main__":
    main()
