import pandas as pd
import os
from pathlib import Path

def split_csv(input_file: str, output_dir: str, chunk_size: int = 50):
    """
    将一个大型CSV文件分割成多个小文件。

    Args:
        input_file (str): 输入的CSV文件路径。
        output_dir (str): 输出分块文件的目录。
        chunk_size (int, optional): 每个分块文件包含的行数。默认为 50。
    """
    # 确保输出目录存在
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # 读取CSV文件
    try:
        df = pd.read_csv(input_file)
        print(f"成功读取CSV文件: {input_file}, 共 {len(df)} 行。")
    except Exception as e:
        print(f"读取CSV文件失败: {e}")
        return

    # 获取表头
    header = df.columns.tolist()
    
    # 计算分块数量
    num_chunks = (len(df) - 1) // chunk_size + 1

    print(f"计划将文件分割成 {num_chunks} 个小文件，每个文件最多 {chunk_size} 行。")

    # 循环创建分块文件
    for i in range(num_chunks):
        start_row = i * chunk_size
        end_row = start_row + chunk_size
        chunk_df = df.iloc[start_row:end_row]
        
        output_filename = Path(output_dir) / f"chunk_{i+1:04d}.csv"
        
        try:
            # 将分块数据写入新的CSV文件，包含表头
            chunk_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            if i % 10 == 0 or i == num_chunks -1:
                 print(f"成功创建分块文件: {output_filename}")
        except Exception as e:
            print(f"写入分块文件 {output_filename} 失败: {e}")

if __name__ == "__main__":
    # 定义输入文件和输出目录
    INPUT_CSV_PATH = r"d:\桌面\clientab-main\temp-data\房地产周刊排版11月4日刊.csv"
    OUTPUT_CHUNKS_DIR = r"d:\桌面\clientab-main\temp-data\chunks"
    
    # 执行分割
    split_csv(INPUT_CSV_PATH, OUTPUT_CHUNKS_DIR)