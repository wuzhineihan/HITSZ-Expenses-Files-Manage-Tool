import pandas as pd
import os
import json
from pathlib import Path
import hashlib
import sys

# 设置标准输出编码为UTF-8,解决Windows命令行emoji显示问题
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 读取 Excel 文件
excel_file = '社团报销.xlsx'
df = pd.read_excel(excel_file)

# 元数据文件路径
metadata_file = 'folder_metadata.json'

# 确保DataFrame有唯一ID列
if '唯一ID' not in df.columns:
    # 如果没有唯一ID列,添加一个
    df['唯一ID'] = None
    print("✨ 添加'唯一ID'列到Excel")

# 确保DataFrame有文件数量列
if '文件数量' not in df.columns:
    # 如果没有文件数量列,添加一个
    df['文件数量'] = None
    print("✨ 添加'文件数量'列到Excel")

# 生成唯一ID的函数
def generate_unique_id():
    """生成一个基于时间戳和随机数的唯一ID"""
    import time
    import random
    timestamp = str(int(time.time() * 1000))
    random_num = str(random.randint(1000, 9999))
    return f"{timestamp}_{random_num}"

def find_matching_metadata(payer, content, metadata):
    """通过付款人和开票内容查找匹配的元数据"""
    for uid, meta in metadata.items():
        # 检查付款人是否匹配
        if meta.get('original_payer') == payer:
            # 检查开票内容是否相似(支持模糊匹配)
            if meta.get('original_content') == content or meta.get('current_content') == content:
                return uid
    return None

# 加载元数据
def load_metadata():
    """加载现有的元数据文件"""
    if os.path.exists(metadata_file):
        try:
            with open(metadata_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️  加载元数据失败: {e}")
            return {}
    return {}

# 保存元数据
def save_metadata(metadata):
    """保存元数据到文件"""
    try:
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, indent=2, ensure_ascii=False, fp=f)
        print(f"✅ 元数据已保存到: {metadata_file}")
    except Exception as e:
        print(f"❌ 保存元数据失败: {e}")

# 加载现有元数据
metadata = load_metadata()

# 获取当前工作目录
base_dir = Path('.')

# 定义状态文件夹
completed_dir = base_dir / '✅已完成'
pending_dir = base_dir / '📋待处理'

# 确保状态文件夹存在
completed_dir.mkdir(exist_ok=True)
pending_dir.mkdir(exist_ok=True)

print("📂 文件夹分类说明:")
print(f"   ✅已完成: 材料准备状态为'yes'的文件夹")
print(f"   📋待处理: 材料准备状态不为'yes'的文件夹")
print()

# 记录Excel中使用的唯一ID
active_unique_ids = set()

# 统计计数器
stats = {
    'completed': 0,
    'pending': 0,
    'moved': 0,
    'created': 0
}

# 遍历每一行(跳过第一行标题)
for index, row in df.iterrows():
    # 获取付款人和开票内容
    # 假设列名为"付款人"和"开票内容",如果列名不同需要调整
    payer = row.get('付款人')
    content = row.get('开票内容')
    unique_id = row.get('唯一ID')
    
    # 如果付款人或开票内容为空,跳过这一行
    if pd.isna(payer) or pd.isna(content) or str(payer).strip() == '' or str(content).strip() == '':
        print(f"跳过第 {index + 2} 行: 付款人或开票内容为空")
        continue
    
    # 清理字符串,去除前后空格
    payer = str(payer).strip()
    content = str(content).strip()
    
    # 计算Excel中的行号(从2开始,因为第1行是标题)
    excel_row_number = index + 2
    
    # 为开票内容添加行号前缀
    content_with_prefix = f"{excel_row_number}.{content}"
    
    # 处理唯一ID
    if pd.isna(unique_id) or str(unique_id).strip() == '':
        # 尝试通过付款人和内容查找现有的元数据
        matched_uid = find_matching_metadata(payer, content, metadata)
        
        if matched_uid:
            # 找到匹配的元数据,重用这个ID
            unique_id = matched_uid
            df.at[index, '唯一ID'] = unique_id
            print(f"🔗 第 {index + 2} 行找到匹配的记录,使用ID: {unique_id}")
        else:
            # 生成新的唯一ID
            unique_id = generate_unique_id()
            df.at[index, '唯一ID'] = unique_id
            print(f"✨ 第 {index + 2} 行生成新ID: {unique_id}")
    else:
        unique_id = str(unique_id).strip()
        print(f"📌 第 {index + 2} 行使用现有ID: {unique_id}")
    
    # 记录活跃的ID
    active_unique_ids.add(unique_id)
    
    # 使用唯一ID作为key
    row_key = unique_id
    
    # 获取当前"材料准备"列的值,决定放在哪个顶级目录
    current_status = row.get('材料准备')
    if current_status == 'yes':
        status_dir = completed_dir
        stats['completed'] += 1
    else:
        status_dir = pending_dir
        stats['pending'] += 1
    
    if row_key in metadata:
        # 已存在元数据,检查是否需要重命名/移动文件夹
        old_folder_path = Path(metadata[row_key]['folder_path'])
        new_folder_path = status_dir / payer / content_with_prefix
        
        if old_folder_path != new_folder_path:
            # 需要重命名/移动文件夹
            if old_folder_path.exists():
                try:
                    # 确保新路径的父目录存在
                    new_folder_path.parent.mkdir(parents=True, exist_ok=True)
                    # 移动文件夹(保留所有文件)
                    import shutil
                    shutil.move(str(old_folder_path), str(new_folder_path))
                    print(f"📦 移动文件夹:")
                    print(f"   从: {old_folder_path}")
                    print(f"   到: {new_folder_path}")
                    folder_path = new_folder_path
                    stats['moved'] += 1
                    
                    # 清理可能为空的旧父文件夹
                    try:
                        old_parent = old_folder_path.parent
                        if old_parent.exists() and old_parent != base_dir and old_parent not in [completed_dir, pending_dir]:
                            if not any(old_parent.iterdir()):
                                old_parent.rmdir()
                                print(f"   🧹 清理空文件夹: {old_parent}")
                    except:
                        pass
                except Exception as e:
                    print(f"❌ 移动文件夹失败: {e}")
                    print(f"   将创建新文件夹: {new_folder_path}")
                    folder_path = new_folder_path
                    folder_path.mkdir(parents=True, exist_ok=True)
            else:
                # 旧文件夹不存在,创建新文件夹
                print(f"⚠️  旧文件夹不存在: {old_folder_path}")
                print(f"   创建新文件夹: {new_folder_path}")
                folder_path = new_folder_path
                folder_path.mkdir(parents=True, exist_ok=True)
        else:
            # 路径没有变化,使用现有文件夹
            folder_path = new_folder_path
            if not folder_path.exists():
                folder_path.mkdir(parents=True, exist_ok=True)
                print(f"创建文件夹: {folder_path}")
            else:
                print(f"使用现有文件夹: {folder_path}")
    else:
        # 新行,创建文件夹和元数据
        folder_path = status_dir / payer / content_with_prefix
        
        # 创建文件夹(如果不存在)
        try:
            folder_path.mkdir(parents=True, exist_ok=True)
            print(f"✨ 创建新文件夹: {folder_path}")
            stats['created'] += 1
        except Exception as e:
            print(f"❌ 创建文件夹失败 {folder_path}: {e}")
            continue
        
        # 创建新的元数据条目
        metadata[row_key] = {
            'unique_id': unique_id,
            'original_payer': payer,
            'original_content': content,
            'original_content_with_prefix': content_with_prefix,
            'folder_path': str(folder_path),
            'created_at': pd.Timestamp.now().isoformat(),
            'excel_row': excel_row_number
        }
    
    # 更新元数据中的当前信息
    metadata[row_key]['current_payer'] = payer
    metadata[row_key]['current_content'] = content
    metadata[row_key]['current_content_with_prefix'] = content_with_prefix
    metadata[row_key]['folder_path'] = str(folder_path)
    metadata[row_key]['last_updated'] = pd.Timestamp.now().isoformat()
    metadata[row_key]['excel_row'] = excel_row_number
    
    # 检测文件夹中的文件数量
    try:
        # 获取文件夹中的所有文件(不包括子文件夹)
        files = [f for f in folder_path.iterdir() if f.is_file()]
        file_count = len(files)
        
        # 将文件数量写入"文件数量"列
        df.at[index, '文件数量'] = file_count
        
        print(f"  - 文件夹 {folder_path} 中有 {file_count} 个文件 (当前状态: {current_status})")
        
        # 如果文件数量小于3,将"材料准备"列设置为"no"
        if file_count < 3:
            df.at[index, '材料准备'] = 'no'
            print(f"  - 文件数量不足3个,将材料准备列设置为 no")
        # 如果文件数量>=3,且材料准备列不为yes,将其设置为"check"
        elif file_count >= 3 and current_status != 'yes':
            df.at[index, '材料准备'] = 'check'
            print(f"  - 文件数量>=3且状态不为yes,将材料准备列设置为 check")
        # 如果文件数量>=3且状态为yes,保持不变
        elif file_count >= 3 and current_status == 'yes':
            print(f"  - 文件数量>=3且状态为yes,保持不变")
    except Exception as e:
        print(f"  - 检查文件夹失败 {folder_path}: {e}")
        # 如果检查失败,将文件数量设置为0
        df.at[index, '文件数量'] = 0

print("\n✅ 所有文件夹处理完成!")
print(f"\n📊 统计信息:")
print(f"   ✅已完成: {stats['completed']} 个文件夹")
print(f"   📋待处理: {stats['pending']} 个文件夹")
print(f"   📦移动: {stats['moved']} 个文件夹")
print(f"   ✨新建: {stats['created']} 个文件夹")

# 清理旧格式的元数据(使用数字索引作为key的旧记录)
old_format_keys = [k for k in metadata.keys() if k.isdigit()]
if old_format_keys:
    print(f"\n🔄 检测到 {len(old_format_keys)} 个旧格式的元数据记录,正在清理...")
    for old_key in old_format_keys:
        # 移除旧格式的记录
        del metadata[old_key]
    print(f"✅ 已清理旧格式记录")

# 检查是否有被删除的行(元数据中存在但Excel中不存在的ID)
orphaned_ids = set(metadata.keys()) - active_unique_ids
if orphaned_ids:
    print(f"\n🗑️  检测到 {len(orphaned_ids)} 个已删除的记录:")
    for orphaned_id in orphaned_ids:
        orphaned_meta = metadata[orphaned_id]
        orphaned_folder = Path(orphaned_meta['folder_path'])
        
        # 兼容旧格式(payer)和新格式(original_payer)
        payer_name = orphaned_meta.get('original_payer') or orphaned_meta.get('payer') or orphaned_meta.get('current_payer')
        content_name = orphaned_meta.get('original_content') or orphaned_meta.get('current_content')
        
        print(f"\n   ID: {orphaned_id}")
        print(f"   付款人: {payer_name}")
        print(f"   开票内容: {content_name}")
        print(f"   文件夹: {orphaned_folder}")
        
        if orphaned_folder.exists():
            # 检查文件夹中是否有文件
            files = [f for f in orphaned_folder.iterdir() if f.is_file()]
            if files:
                print(f"   ⚠️  文件夹中还有 {len(files)} 个文件,已保留")
            else:
                # 文件夹为空,删除它
                try:
                    orphaned_folder.rmdir()
                    print(f"   ✅ 文件夹为空,已删除")
                    
                    # 检查父文件夹(付款人文件夹)是否也为空
                    parent_folder = orphaned_folder.parent
                    if parent_folder != base_dir and parent_folder.exists() and parent_folder not in [completed_dir, pending_dir]:
                        # 检查父文件夹是否为空
                        try:
                            if not any(parent_folder.iterdir()):
                                parent_folder.rmdir()
                                print(f"   ✅ 父文件夹 {parent_folder.name} 也为空,已删除")
                                
                                # 检查祖父文件夹(状态文件夹下的空付款人文件夹)
                                grandparent_folder = parent_folder.parent
                                if grandparent_folder in [completed_dir, pending_dir] and grandparent_folder.exists():
                                    try:
                                        if not any(grandparent_folder.iterdir()):
                                            pass  # 不删除✅已完成和📋待处理文件夹本身
                                    except:
                                        pass
                        except Exception as e:
                            pass  # 父文件夹不为空或删除失败,忽略
                except Exception as e:
                    print(f"   ❌ 删除空文件夹失败: {e}")
        else:
            print(f"   ℹ️  文件夹不存在")
        
        # 标记为已删除(保留元数据以便恢复)
        metadata[orphaned_id]['deleted'] = True
        metadata[orphaned_id]['deleted_at'] = pd.Timestamp.now().isoformat()

# 收集所有活跃的文件夹路径
active_folder_paths = set()
for uid in active_unique_ids:
    if uid in metadata:
        active_folder_paths.add(Path(metadata[uid]['folder_path']))

# 清理不在活跃列表中的所有文件夹
orphaned_folders_cleaned = []
for status_folder in [completed_dir, pending_dir]:
    if status_folder.exists():
        for payer_folder in status_folder.iterdir():
            if payer_folder.is_dir():
                # 遍历付款人文件夹下的所有文件夹
                for content_folder in payer_folder.iterdir():
                    if content_folder.is_dir():
                        # 检查这个文件夹是否在活跃列表中
                        if content_folder not in active_folder_paths:
                            # 不在活跃列表中,检查是否有文件
                            files = [f for f in content_folder.rglob('*') if f.is_file()]
                            if files:
                                print(f"\n⚠️  发现未追踪的文件夹(有文件,已保留):")
                                print(f"   路径: {content_folder}")
                                print(f"   文件数: {len(files)} 个")
                            else:
                                # 空文件夹,删除它
                                try:
                                    import shutil
                                    shutil.rmtree(content_folder)
                                    orphaned_folders_cleaned.append(str(content_folder))
                                    print(f"\n🧹 清理未追踪的空文件夹: {content_folder}")
                                except Exception as e:
                                    print(f"\n❌ 清理文件夹失败 {content_folder}: {e}")

# 清理空的付款人文件夹(在✅已完成和📋待处理文件夹中)
empty_folders_deleted = []
for status_folder in [completed_dir, pending_dir]:
    if status_folder.exists():
        for item in status_folder.iterdir():
            if item.is_dir():
                # 检查是否为空文件夹
                try:
                    if not any(item.iterdir()):
                        item.rmdir()
                        empty_folders_deleted.append(f"{status_folder.name}/{item.name}")
                except Exception as e:
                    pass  # 忽略错误

# 同时清理根目录下的旧文件夹(不在状态文件夹中的)
for item in base_dir.iterdir():
    if item.is_dir() and item.name not in ['.git', '.azure', '✅已完成', '📋待处理', '.venv', '__pycache__']:
        # 检查是否为空文件夹
        try:
            if not any(item.iterdir()):
                item.rmdir()
                empty_folders_deleted.append(item.name)
        except Exception as e:
            pass  # 忽略错误

if empty_folders_deleted:
    print(f"\n🧹 清理空的付款人文件夹:")
    for folder_name in empty_folders_deleted:
        print(f"   ✅ 已删除: {folder_name}")

# 保存元数据
save_metadata(metadata)

# 保存更新后的 Excel 文件
try:
    df.to_excel(excel_file, index=False)
    print(f"\nExcel 文件已更新: {excel_file}")
except PermissionError:
    print(f"\n⚠️  无法保存 Excel 文件: {excel_file}")
    print("可能的原因:")
    print("1. Excel 文件正在被 Microsoft Excel 或其他程序打开")
    print("2. 文件被设置为只读")
    print("\n解决方法:")
    print("1. 关闭 Excel 文件后重新运行此脚本")
    print("2. 或者脚本会创建一个备份文件: 社团报销_updated.xlsx")
    
    # 尝试保存为新文件
    try:
        backup_file = '社团报销_updated.xlsx'
        df.to_excel(backup_file, index=False)
        print(f"\n✅ 已保存为新文件: {backup_file}")
    except Exception as e:
        print(f"\n❌ 保存备份文件也失败: {e}")
except Exception as e:
    print(f"\n保存 Excel 文件失败: {e}")
