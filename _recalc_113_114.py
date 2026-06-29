#!/usr/bin/env python3
"""
用Excel COM重算113/114批报价单的公式缓存
"""
import os
import sys

# 使用系统Python（有pywin32）
sys.path.insert(0, 'D:/baojia/electron/app')

try:
    import win32com.client
    print("✅ pywin32 可用")
except ImportError:
    print("❌ pywin32 不可用，请使用系统Python运行")
    sys.exit(1)

# 文件列表
base = r'\\192.168.110.111\【管理者专用】\04合同\000-DQ业务对接用\报价结算系统\报价\叠纸\ニキ新作\已结算\2026年6月'
targets = [
    '113批报价单_ニキ新作_中译日.xlsx',
    '114批报价单_ニキ新作_中译日.xlsx',
]

print(f"\n基类路径: {base}")
print(f"目标文件: {targets}\n")

# 启动Excel
print("正在启动Excel...")
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False  # 后台运行
excel.DisplayAlerts = False  # 不显示警告
print("✅ Excel已启动\n")

for name in targets:
    fpath = os.path.join(base, name)
    print(f"处理: {name}")
    
    if not os.path.exists(fpath):
        print(f"  ❌ 文件不存在: {fpath}")
        continue
    
    try:
        # 打开文件
        print(f"  打开文件...")
        wb = excel.Workbooks.Open(fpath)
        
        # 强制重算所有公式
        print(f"  重算公式...")
        excel.CalculateUntilAsyncQueriesDone()
        
        # 验证G48/I48有值
        ws = wb.ActiveSheet
        g48 = ws.Cells(48, 7).Value  # G列=7
        i48 = ws.Cells(48, 9).Value  # I列=9
        print(f"  G48={g48}, I48={i48}")
        
        # 保存
        print(f"  保存文件...")
        wb.Save()
        wb.Close()
        print(f"  ✅ 完成\n")
        
    except Exception as e:
        print(f"  ❌ 失败: {e}\n")

# 关闭Excel
excel.Quit()
print("✅ Excel已关闭")
print("\n全部完成！现在可以重新测试结算预览了。")
