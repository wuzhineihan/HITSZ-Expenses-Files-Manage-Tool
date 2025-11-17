"""
文件监控脚本 - 自动检测Excel文件变化并运行创建文件夹脚本
当社团报销.xlsx被修改保存后,自动执行create_folders.py
"""

import time
import os
import subprocess
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ExcelFileHandler(FileSystemEventHandler):
    """监控Excel文件变化的处理器"""
    
    def __init__(self, excel_file, script_file):
        self.excel_file = Path(excel_file).resolve()
        self.script_file = Path(script_file).resolve()
        self.last_modified = 0
        self.cooldown = 2  # 冷却时间(秒),避免重复触发
        self.pending_execution = False  # 是否有待执行的任务
        
    def is_file_locked(self, filepath):
        """检查文件是否被占用(被Excel打开)"""
        try:
            # 尝试以独占模式打开文件
            with open(filepath, 'r+b') as f:
                pass
            return False  # 文件未被占用
        except (IOError, PermissionError):
            return True  # 文件被占用
            
    def wait_for_file_close(self, filepath, max_wait=30):
        """等待文件被关闭,最多等待max_wait秒"""
        print(f"⏳ 检测到Excel文件正在使用中,等待文件关闭...")
        print(f"💡 提示: 请在Excel中关闭文件后,脚本会自动执行")
        
        wait_time = 0
        check_interval = 1  # 每秒检查一次
        
        while wait_time < max_wait:
            time.sleep(check_interval)
            wait_time += check_interval
            
            if not self.is_file_locked(filepath):
                print(f"✅ 文件已关闭 (等待了 {wait_time} 秒)")
                return True
            
            # 每5秒显示一次等待提示
            if wait_time % 5 == 0:
                print(f"⏳ 仍在等待... ({wait_time}/{max_wait}秒)")
        
        print(f"⚠️ 等待超时 ({max_wait}秒),文件仍被占用")
        print(f"💡 请手动关闭Excel文件后,再次保存以触发脚本")
        return False
        
    def on_modified(self, event):
        """文件被修改时触发"""
        if event.is_directory:
            return
            
        file_path = Path(event.src_path).resolve()
        
        # 检查是否是目标Excel文件
        if file_path == self.excel_file:
            current_time = time.time()
            
            # 防止短时间内重复触发
            if current_time - self.last_modified < self.cooldown:
                return
                
            self.last_modified = current_time
            
            print(f"\n{'='*60}")
            print(f"📝 检测到Excel文件变化: {file_path.name}")
            print(f"⏰ 时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*60}\n")
            
            # 等待一小段时间,确保文件已保存完成
            time.sleep(0.5)
            
            # 检查文件是否被占用(Excel是否已关闭)
            if self.is_file_locked(file_path):
                # 文件被占用,等待关闭
                if not self.wait_for_file_close(file_path, max_wait=60):
                    print("⏭️  跳过本次执行,等待下次文件保存\n")
                    print(f"{'='*60}")
                    print("👀 继续监控文件变化...\n")
                    return
            
            # 文件已关闭或从未被打开,可以安全执行脚本
            try:
                print("🚀 正在执行文件夹创建脚本...\n")
                result = subprocess.run(
                    ['python', str(self.script_file)],
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='ignore'
                )
                
                # 显示输出
                if result.stdout:
                    print(result.stdout)
                    
                if result.stderr:
                    print("⚠️ 错误信息:")
                    print(result.stderr)
                    
                if result.returncode == 0:
                    print("\n✅ 执行完成!\n")
                else:
                    print(f"\n❌ 执行失败,退出码: {result.returncode}\n")
                    
            except Exception as e:
                print(f"❌ 执行脚本时出错: {e}\n")
            
            print(f"{'='*60}")
            print("👀 继续监控文件变化...\n")

def main():
    # 配置
    excel_file = '社团报销.xlsx'
    script_file = 'create_folders.py'
    watch_dir = Path('.').resolve()
    
    # 检查文件是否存在
    if not Path(excel_file).exists():
        print(f"❌ 错误: 找不到Excel文件 '{excel_file}'")
        return
        
    if not Path(script_file).exists():
        print(f"❌ 错误: 找不到脚本文件 '{script_file}'")
        return
    
    print("="*60)
    print("📂 社团报销自动化监控系统")
    print("="*60)
    print(f"📁 监控目录: {watch_dir}")
    print(f"📊 监控文件: {excel_file}")
    print(f"🔧 执行脚本: {script_file}")
    print("="*60)
    print("\n✅ 监控已启动!")
    print("💡 提示: 每次保存Excel文件后会自动执行脚本")
    print("⚠️  按 Ctrl+C 可以停止监控\n")
    
    # 创建事件处理器和观察者
    event_handler = ExcelFileHandler(excel_file, script_file)
    observer = Observer()
    observer.schedule(event_handler, str(watch_dir), recursive=False)
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n⏹️  停止监控...")
        observer.stop()
        
    observer.join()
    print("✅ 监控已停止\n")

if __name__ == "__main__":
    main()
