import pandas as pd
import os
import json
from pathlib import Path
import hashlib

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

# 记录Excel中使用的唯一ID
active_unique_ids = set()

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
    
    if row_key in metadata:
        # 已存在元数据,检查是否需要重命名文件夹
        old_folder_path = Path(metadata[row_key]['folder_path'])
        new_folder_path = base_dir / payer / content
        
        if old_folder_path != new_folder_path:
            # 需要重命名文件夹
            if old_folder_path.exists():
                try:
                    # 确保新路径的父目录存在
                    new_folder_path.parent.mkdir(parents=True, exist_ok=True)
                    # 重命名文件夹
                    old_folder_path.rename(new_folder_path)
                    print(f"📝 重命名文件夹:")
                    print(f"   从: {old_folder_path}")
                    print(f"   到: {new_folder_path}")
                    folder_path = new_folder_path
                except Exception as e:
                    print(f"❌ 重命名文件夹失败: {e}")
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
        folder_path = base_dir / payer / content
        
        # 创建文件夹(如果不存在)
        try:
            folder_path.mkdir(parents=True, exist_ok=True)
            print(f"✨ 创建新文件夹: {folder_path}")
        except Exception as e:
            print(f"❌ 创建文件夹失败 {folder_path}: {e}")
            continue
        
        # 创建新的元数据条目
        metadata[row_key] = {
            'unique_id': unique_id,
            'original_payer': payer,
            'original_content': content,
            'folder_path': str(folder_path),
            'created_at': pd.Timestamp.now().isoformat()
        }
    
    # 更新元数据中的当前信息
    metadata[row_key]['current_payer'] = payer
    metadata[row_key]['current_content'] = content
    metadata[row_key]['folder_path'] = str(folder_path)
    metadata[row_key]['last_updated'] = pd.Timestamp.now().isoformat()
    
    # 检测文件夹中的文件数量
    try:
        # 获取文件夹中的所有文件(不包括子文件夹)
        files = [f for f in folder_path.iterdir() if f.is_file()]
        file_count = len(files)
        
        # 获取当前"材料准备"列的值
        current_status = row.get('材料准备')
        
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

print("\n所有文件夹创建完成!")

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
                print(f"   ℹ️  文件夹为空,已保留(可手动删除)")
        else:
            print(f"   ℹ️  文件夹不存在")
        
        # 标记为已删除(保留元数据以便恢复)
        metadata[orphaned_id]['deleted'] = True
        metadata[orphaned_id]['deleted_at'] = pd.Timestamp.now().isoformat()

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
