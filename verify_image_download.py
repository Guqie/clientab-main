#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片下载功能验证脚本
用于测试集成后的新浪图片下载策略是否正常工作
"""

import os
import sys
import logging
from pathlib import Path

# 添加项目路径到sys.path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / 'src'))

from csv_word_converter.core import csv_to_word_universal, UniversalDocumentGenerator

def setup_logging():
    """
    设置详细的日志记录，用于调试图片下载问题
    """
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('image_download_debug.log', encoding='utf-8')
        ]
    )
    return logging.getLogger(__name__)

def test_direct_image_download():
    """
    直接测试UniversalDocumentGenerator的图片下载功能
    """
    logger = logging.getLogger(__name__)
    logger.info("=== 开始直接图片下载测试 ===")
    
    try:
        # 创建生成器实例
        generator = UniversalDocumentGenerator('new_energy', 'templates_config.yaml')
        
        # 测试新浪图片URL
        test_urls = [
            "https://n.sinaimg.cn/finance/crawl/200/w600h400/20200101/abcd-1234567.jpg",
            "https://k.sinaimg.cn/n/finance/transform/266/w400h266/20200102/efgh-2345678.jpg",
            "https://n.sinaimg.cn/spider/transform/266/w400h266/20200103/ijkl-3456789.jpg"
        ]
        
        success_count = 0
        for i, url in enumerate(test_urls, 1):
            logger.info(f"测试图片 {i}: {url}")
            try:
                # 直接调用_replace_url方法测试图片下载
                result = generator._replace_url(url, url)
                if result != url:  # 如果返回值不是原URL，说明下载成功
                    logger.info(f"✅ 图片 {i} 下载成功: {result}")
                    success_count += 1
                else:
                    logger.warning(f"❌ 图片 {i} 下载失败，返回原URL")
            except Exception as e:
                logger.error(f"❌ 图片 {i} 下载异常: {str(e)}")
        
        logger.info(f"直接下载测试完成，成功率: {success_count}/{len(test_urls)} ({success_count/len(test_urls)*100:.1f}%)")
        return success_count > 0
        
    except Exception as e:
        logger.error(f"直接图片下载测试失败: {str(e)}")
        return False

def test_csv_to_word_conversion():
    """
    测试完整的CSV转Word流程，包含图片下载
    """
    logger = logging.getLogger(__name__)
    logger.info("=== 开始CSV转Word集成测试 ===")
    
    try:
        # 检查测试文件是否存在
        csv_file = "test_sina_images.csv"
        config_file = "templates_config.yaml"
        
        if not os.path.exists(csv_file):
            logger.error(f"测试CSV文件不存在: {csv_file}")
            return False
            
        if not os.path.exists(config_file):
            logger.error(f"配置文件不存在: {config_file}")
            return False
        
        logger.info(f"使用CSV文件: {csv_file}")
        logger.info(f"使用配置文件: {config_file}")
        logger.info(f"使用模板类型: new_energy")
        
        # 调用csv_to_word_universal函数（使用正确的参数）
        output_path = csv_to_word_universal(
            csv_file=csv_file,
            template_type='new_energy',
            config_path=config_file
        )
        
        if output_path and os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            logger.info(f"✅ CSV转Word成功完成!")
            logger.info(f"输出文件: {output_path}")
            logger.info(f"文件大小: {file_size:,} 字节")
            
            # 检查temp-images目录是否有下载的图片
            temp_images_dir = "temp-images"
            if os.path.exists(temp_images_dir):
                image_files = [f for f in os.listdir(temp_images_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                logger.info(f"下载的图片文件数量: {len(image_files)}")
                for img_file in image_files:
                    img_path = os.path.join(temp_images_dir, img_file)
                    img_size = os.path.getsize(img_path)
                    logger.info(f"  - {img_file}: {img_size:,} 字节")
            else:
                logger.warning("temp-images目录不存在，可能没有下载任何图片")
            
            return True
        else:
            logger.error("❌ CSV转Word失败，没有生成输出文件")
            return False
            
    except Exception as e:
        logger.error(f"❌ CSV转Word集成测试失败: {str(e)}")
        import traceback
        logger.error(f"详细错误信息:\n{traceback.format_exc()}")
        return False

def main():
    """
    主函数：运行所有测试
    """
    logger = setup_logging()
    logger.info("开始图片下载功能验证测试")
    
    # 确保在正确的工作目录
    os.chdir(project_root)
    logger.info(f"工作目录: {os.getcwd()}")
    
    # 测试结果统计
    results = {
        "直接图片下载测试": False,
        "CSV转Word集成测试": False
    }
    
    # 运行测试
    try:
        results["直接图片下载测试"] = test_direct_image_download()
        results["CSV转Word集成测试"] = test_csv_to_word_conversion()
    except Exception as e:
        logger.error(f"测试运行异常: {str(e)}")
    
    # 输出测试结果汇总
    logger.info("\n" + "="*50)
    logger.info("测试结果汇总:")
    for test_name, result in results.items():
        status = "✅ 通过" if result else "❌ 失败"
        logger.info(f"  {test_name}: {status}")
    
    success_count = sum(results.values())
    total_count = len(results)
    success_rate = success_count / total_count * 100
    
    logger.info(f"\n总体成功率: {success_count}/{total_count} ({success_rate:.1f}%)")
    
    if success_rate < 100:
        logger.warning("\n⚠️  存在失败的测试，请检查日志文件 image_download_debug.log 获取详细信息")
        logger.warning("建议检查以下方面:")
        logger.warning("1. 网络连接是否正常")
        logger.warning("2. 新浪图片URL是否有效")
        logger.warning("3. 请求头配置是否正确")
        logger.warning("4. 图片验证逻辑是否过于严格")
    else:
        logger.info("\n🎉 所有测试通过！图片下载功能正常工作。")

if __name__ == "__main__":
    main()