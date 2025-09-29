# -*- coding: utf-8 -*-
"""
内容解析器模块
==============================================================================

本模块负责将结构化数据转换为Word文档内容，处理不同期刊的内容结构和格式要求。

主要功能:
1. 内容结构解析与组织
2. 标题层级管理
3. 文章格式化处理
4. 特殊内容处理（如加粗、链接等）

"""

import re
import logging
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from enum import Enum

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ContentType(Enum):
    """内容类型枚举"""
    TITLE_H1 = "h1"  # 一级标题
    TITLE_H2 = "h2"  # 二级标题
    TITLE_H3 = "h3"  # 三级标题
    TITLE_H4 = "h4"  # 四级标题
    ARTICLE = "article"  # 文章内容
    SOURCE = "source"  # 来源信息
    NORMAL = "normal"  # 普通文本
    IMAGE = "image"  # 图片内容

@dataclass
class ContentBlock:
    """内容块数据结构"""
    content_type: ContentType
    text: str
    level: int = 0  # 大纲级别
    metadata: Dict[str, Any] = None
    image_path: str = None  # 图片路径（用于IMAGE类型）
    
    def __post_init__(self):
        if self.metadata is None:
            self.metadata = {}

class WeeklyContentParser:
    """
    周刊内容解析器
    
    负责将结构化数据转换为Word文档内容块，支持不同期刊的内容结构。
    """
    
    def __init__(self, weekly_type: str = "real_estate_weekly"):
        """
        初始化内容解析器
        
        Args:
            weekly_type (str): 周刊类型
        """
        self.weekly_type = weekly_type
        self.content_blocks = []
        
        # 根据周刊类型设置内容结构
        self._setup_content_structure()
    
    def _setup_content_structure(self):
        """
        根据周刊类型设置内容结构
        """
        if self.weekly_type == "real_estate_weekly":
            self._setup_real_estate_structure()
        elif self.weekly_type == "electricity_weekly":
            self._setup_electricity_structure()
        else:
            self._setup_default_structure()
    
    def _setup_real_estate_structure(self):
        """
        设置房地产周刊内容结构
        """
        self.category_mapping = {
            "政策环境": {
                "level": 1,
                "subcategories": {
                    "国家政策": {"level": 2},
                    "地方政策": {"level": 2}
                }
            },
            "市场动态": {
                "level": 1,
                "subcategories": {
                    "土地市场": {"level": 2},
                    "房价走势": {"level": 2},
                    "成交数据": {"level": 2}
                }
            },
            "企业动向": {
                "level": 1,
                "subcategories": {
                    "房企动态": {"level": 2},
                    "投资并购": {"level": 2}
                }
            }
        }
    
    def _setup_electricity_structure(self):
        """
        设置电力周刊内容结构
        """
        self.category_mapping = {
            "政策环境": {
                "level": 1,
                "subcategories": {
                    "国家政策": {"level": 2},
                    "地方政策": {"level": 2}
                }
            },
            "行业要讯": {
                "level": 1,
                "subcategories": {
                    "行业动态": {"level": 2},
                    "技术发展": {"level": 2}
                }
            },
            "竞争动向": {
                "level": 1,
                "subcategories": {
                    "五大发电集团": {"level": 2},
                    "其他发电集团": {"level": 2},
                    "输配电集团": {"level": 2},
                    "辅业集团": {"level": 2},
                    "电力设备": {"level": 2}
                }
            },
            "国际要讯": {
                "level": 1,
                "subcategories": {}
            },
            "产业链分析": {
                "level": 1,
                "subcategories": {}
            },
            "案例借鉴": {
                "level": 1,
                "subcategories": {}
            }
        }
        
        # 添加CSV格式到标准格式的映射
        self.csv_category_mapping = {
            "一、政策环境": "政策环境",
            "二、行业要讯": "行业要讯", 
            "三、竞争动向": "竞争动向",
            "四、国际要讯": "国际要讯",
            "五、产业分析": "产业链分析",
            "六、案例借鉴": "案例借鉴"
        }
    
    def _setup_default_structure(self):
        """
        设置默认内容结构
        """
        self.category_mapping = {
            "默认分类": {
                "level": 1,
                "subcategories": {
                    "默认子分类": {"level": 2}
                }
            }
        }
        
        # 添加默认的CSV格式映射
        self.csv_category_mapping = {}
    
    def parse_grouped_data(self, grouped_data: Dict[str, Dict[str, List[Dict]]]) -> List[ContentBlock]:
        """
        解析分组数据为内容块
        
        Args:
            grouped_data (Dict): 分组后的数据
            
        Returns:
            List[ContentBlock]: 内容块列表
        """
        self.content_blocks = []
        
        # 创建映射后的数据结构
        mapped_data = {}
        
        # 处理CSV格式的分类名称映射和数据合并
        for csv_category, subcategories in grouped_data.items():
            # 跳过无效分类（nan等）
            if not csv_category or str(csv_category).lower() in ['nan', 'none', '']:
                # 对于无效分类，需要根据子分类来分配文章
                self._distribute_orphaned_articles(subcategories, mapped_data)
                continue
                
            # 映射CSV格式的分类名称到标准格式
            standard_category = self.csv_category_mapping.get(csv_category, csv_category)
            
            # 如果映射后的分类不在预定义结构中，跳过
            if standard_category not in self.category_mapping:
                continue
                
            # 合并到映射后的数据结构中
            if standard_category not in mapped_data:
                mapped_data[standard_category] = {}
            
            for subcategory, articles in subcategories.items():
                if subcategory not in mapped_data[standard_category]:
                    mapped_data[standard_category][subcategory] = []
                mapped_data[standard_category][subcategory].extend(articles)
        
        # 按预定义顺序处理分类，生成完整框架结构
        for category_name in self.category_mapping.keys():
            self._process_category_with_framework(category_name, mapped_data.get(category_name, {}))
        
        logger.info(f"内容解析完成，共生成 {len(self.content_blocks)} 个内容块")
        return self.content_blocks
    
    def _distribute_orphaned_articles(self, subcategories: Dict[str, List[Dict]], mapped_data: Dict):
        """
        分配没有明确分类的文章到合适的分类中
        
        Args:
            subcategories (Dict): 子分类数据
            mapped_data (Dict): 映射后的数据结构
        """
        for subcategory, articles in subcategories.items():
            # 跳过无效子分类
            if not subcategory or str(subcategory).lower() in ['nan', 'none', '']:
                # 对于完全没有分类信息的文章，暂时放到"行业要讯"分类下
                target_category = "行业要讯"
                target_subcategory = "行业动态"
                
                if target_category not in mapped_data:
                    mapped_data[target_category] = {}
                if target_subcategory not in mapped_data[target_category]:
                    mapped_data[target_category][target_subcategory] = []
                
                mapped_data[target_category][target_subcategory].extend(articles)
                continue
            
            # 根据子分类名称智能分配到合适的主分类
            target_category = self._infer_category_from_subcategory(subcategory)
            
            if target_category not in mapped_data:
                mapped_data[target_category] = {}
            if subcategory not in mapped_data[target_category]:
                mapped_data[target_category][subcategory] = []
            
            mapped_data[target_category][subcategory].extend(articles)
    
    def _infer_category_from_subcategory(self, subcategory: str) -> str:
        """
        根据子分类名称推断主分类
        
        Args:
            subcategory (str): 子分类名称
            
        Returns:
            str: 推断的主分类名称
        """
        # 子分类到主分类的映射规则
        subcategory_mapping = {
            "国家政策": "政策环境",
            "地方政策": "政策环境",
            "行业动态": "行业要讯",
            "技术发展": "行业要讯",
            "五大发电集团": "竞争动向",
            "其他发电集团": "竞争动向",
            "输配电集团": "竞争动向",
            "辅业集团": "竞争动向",
            "电力设备": "竞争动向"
        }
        
        # 直接匹配
        if subcategory in subcategory_mapping:
            return subcategory_mapping[subcategory]
        
        # 关键词匹配
        if any(keyword in subcategory for keyword in ["政策", "法规", "规定"]):
            return "政策环境"
        elif any(keyword in subcategory for keyword in ["国际", "海外", "全球"]):
            return "国际要讯"
        elif any(keyword in subcategory for keyword in ["产业", "链条", "分析"]):
            return "产业链分析"
        elif any(keyword in subcategory for keyword in ["案例", "借鉴", "经验"]):
            return "案例借鉴"
        elif any(keyword in subcategory for keyword in ["竞争", "集团", "公司", "企业"]):
            return "竞争动向"
        else:
            # 默认分类
            return "行业要讯"
    
    def _process_category_with_framework(self, category_name: str, subcategories: Dict[str, List[Dict]]):
        """
        按框架处理分类（包含子分类结构）
        
        Args:
            category_name (str): 分类名称
            subcategories (Dict): 子分类数据
        """
        category_config = self.category_mapping[category_name]
        
        # 始终添加分类标题（一级标题）
        category_title = self._format_category_title(category_name)
        self.content_blocks.append(ContentBlock(
            content_type=ContentType.TITLE_H1,
            text=category_title,
            level=0,
            metadata={"category": category_name}
        ))
        
        # 处理子分类 - 生成完整框架
        subcategory_configs = category_config.get("subcategories", {})
        
        if subcategory_configs:
            # 有预定义的子分类，按预定义顺序生成
            for subcategory_name in subcategory_configs.keys():
                # 始终添加子分类标题（二级标题）
                subcategory_title = self._format_subcategory_title(subcategory_name)
                self.content_blocks.append(ContentBlock(
                    content_type=ContentType.TITLE_H2,
                    text=subcategory_title,
                    level=1,
                    metadata={"category": category_name, "subcategory": subcategory_name}
                ))
                
                # 处理该子分类下的文章（如果有的话）
                articles = subcategories.get(subcategory_name, [])
                for article in articles:
                    self._process_article(article, category_name, subcategory_name)
        else:
            # 没有预定义子分类，直接处理数据中的子分类
            for subcategory_name, articles in subcategories.items():
                if subcategory_name and subcategory_name.strip():
                    subcategory_title = self._format_subcategory_title(subcategory_name)
                    self.content_blocks.append(ContentBlock(
                        content_type=ContentType.TITLE_H2,
                        text=subcategory_title,
                        level=1,
                        metadata={"category": category_name, "subcategory": subcategory_name}
                    ))
                
                # 处理文章
                for article in articles:
                    self._process_article(article, category_name, subcategory_name)
    
    def _process_category(self, category_name: str, subcategories: Dict[str, List[Dict]]):
        """
        处理分类内容（保留原有方法以兼容性）
        
        Args:
            category_name (str): 分类名称
            subcategories (Dict): 子分类数据
        """
        return self._process_category_with_framework(category_name, subcategories)
    
    def _process_unknown_category(self, category_name: str, subcategories: Dict[str, List[Dict]]):
        """
        处理未知分类
        
        Args:
            category_name (str): 分类名称
            subcategories (Dict): 子分类数据
        """
        # 添加分类标题
        category_title = self._format_category_title(category_name)
        self.content_blocks.append(ContentBlock(
            content_type=ContentType.TITLE_H1,
            text=category_title,
            level=0,
            metadata={"category": category_name}
        ))
        
        # 处理子分类和文章
        for subcategory_name, articles in subcategories.items():
            if subcategory_name and subcategory_name.strip():
                subcategory_title = self._format_subcategory_title(subcategory_name)
                self.content_blocks.append(ContentBlock(
                    content_type=ContentType.TITLE_H2,
                    text=subcategory_title,
                    level=1,
                    metadata={"category": category_name, "subcategory": subcategory_name}
                ))
            
            for article in articles:
                self._process_article(article, category_name, subcategory_name)
    
    def _process_article(self, article: Dict, category: str, subcategory: str):
        """
        处理单篇文章
        
        Args:
            article (Dict): 文章数据
            category (str): 分类名称
            subcategory (str): 子分类名称
        """
        title = article.get('标题', '').strip()
        content = article.get('内容', '').strip()
        source = article.get('来源', '').strip()
        date = article.get('日期', '').strip()
        
        if not title:
            return
        
        # 格式化文章标题
        formatted_title = self._format_article_title(title)
        self.content_blocks.append(ContentBlock(
            content_type=ContentType.TITLE_H4,
            text=formatted_title,
            level=3,
            metadata={
                "category": category,
                "subcategory": subcategory,
                "article_type": "title"
            }
        ))
        
        # 添加文章内容
        if content:
            formatted_content = self._format_article_content(content)
            self.content_blocks.append(ContentBlock(
                content_type=ContentType.ARTICLE,
                text=formatted_content,
                level=0,
                metadata={
                    "category": category,
                    "subcategory": subcategory,
                    "article_type": "content"
                }
            ))
        
        # 添加来源信息
        if source or date:
            source_text = self._format_source_info(source, date)
            self.content_blocks.append(ContentBlock(
                content_type=ContentType.SOURCE,
                text=source_text,
                level=0,
                metadata={
                    "category": category,
                    "subcategory": subcategory,
                    "article_type": "source"
                }
            ))
    
    def _format_category_title(self, category_name: str) -> str:
        """
        格式化分类标题
        
        Args:
            category_name (str): 分类名称
            
        Returns:
            str: 格式化后的标题
        """
        # 根据周刊类型添加序号
        if self.weekly_type == "electricity_weekly":
            category_order = list(self.category_mapping.keys())
            if category_name in category_order:
                index = category_order.index(category_name) + 1
                chinese_numbers = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
                if index <= len(chinese_numbers):
                    return f"{chinese_numbers[index-1]}、{category_name}"
        
        return category_name
    
    def _format_subcategory_title(self, subcategory_name: str) -> str:
        """
        格式化子分类标题
        
        Args:
            subcategory_name (str): 子分类名称
            
        Returns:
            str: 格式化后的标题
        """
        # 电力周刊使用〖〗包围
        if self.weekly_type == "electricity_weekly":
            return f"〖{subcategory_name}〗"
        
        return subcategory_name
    
    def _format_article_title(self, title: str) -> str:
        """
        格式化文章标题
        
        Args:
            title (str): 原始标题
            
        Returns:
            str: 格式化后的标题
        """
        # 电力周刊使用【】包围
        if self.weekly_type == "electricity_weekly":
            # 如果已经有【】，则不重复添加
            if not (title.startswith('【') and title.endswith('】')):
                return f"【{title}】"
        
        return title
    
    def _format_article_content(self, content: str) -> str:
        """
        格式化文章内容
        
        Args:
            content (str): 原始内容
            
        Returns:
            str: 格式化后的内容
        """
        # 处理特殊格式
        content = self._process_bold_text(content)
        content = self._process_line_breaks(content)
        
        return content
    
    def _format_source_info(self, source: str, date: str) -> str:
        """
        格式化来源信息
        
        Args:
            source (str): 来源
            date (str): 日期
            
        Returns:
            str: 格式化后的来源信息
        """
        parts = []
        if source:
            parts.append(source)
        if date:
            parts.append(date)
        
        return ' '.join(parts)
    
    def _process_bold_text(self, text: str) -> str:
        """
        处理加粗文本标记
        
        Args:
            text (str): 原始文本
            
        Returns:
            str: 处理后的文本
        """
        # 将**text**标记转换为<bold>text</bold>
        text = re.sub(r'\*\*(.*?)\*\*', r'<bold>\1</bold>', text)
        return text
    
    def _process_line_breaks(self, text: str) -> str:
        """
        处理换行符
        
        Args:
            text (str): 原始文本
            
        Returns:
            str: 处理后的文本
        """
        # 标准化换行符
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # 移除多余的空行
        text = re.sub(r'\n\s*\n\s*\n', '\n\n', text)
        
        return text
    
    def get_content_blocks(self) -> List[ContentBlock]:
        """
        获取内容块列表
        
        Returns:
            List[ContentBlock]: 内容块列表
        """
        return self.content_blocks
    
    def get_statistics(self) -> Dict[str, int]:
        """
        获取内容统计信息
        
        Returns:
            Dict[str, int]: 统计信息
        """
        stats = {
            "total_blocks": len(self.content_blocks),
            "h1_count": 0,
            "h2_count": 0,
            "h3_count": 0,
            "h4_count": 0,
            "article_count": 0,
            "source_count": 0
        }
        
        for block in self.content_blocks:
            if block.content_type == ContentType.TITLE_H1:
                stats["h1_count"] += 1
            elif block.content_type == ContentType.TITLE_H2:
                stats["h2_count"] += 1
            elif block.content_type == ContentType.TITLE_H3:
                stats["h3_count"] += 1
            elif block.content_type == ContentType.TITLE_H4:
                stats["h4_count"] += 1
            elif block.content_type == ContentType.ARTICLE:
                stats["article_count"] += 1
            elif block.content_type == ContentType.SOURCE:
                stats["source_count"] += 1
        
        return stats


# 便捷函数
# ==============================================================================

def parse_weekly_content(grouped_data: Dict[str, Dict[str, List[Dict]]], 
                        weekly_type: str = "real_estate_weekly") -> List[ContentBlock]:
    """
    便捷函数：解析周刊内容
    
    Args:
        grouped_data (Dict): 分组后的数据
        weekly_type (str): 周刊类型
        
    Returns:
        List[ContentBlock]: 内容块列表
    """
    parser = WeeklyContentParser(weekly_type)
    return parser.parse_grouped_data(grouped_data)


if __name__ == "__main__":
    # 测试代码
    test_data = {
        "政策环境": {
            "国家政策": [
                {
                    "标题": "测试标题1",
                    "内容": "测试内容1",
                    "来源": "测试来源",
                    "日期": "2025/01/16"
                }
            ]
        }
    }
    
    parser = WeeklyContentParser("electricity_weekly")
    blocks = parser.parse_grouped_data(test_data)
    
    print(f"生成了 {len(blocks)} 个内容块:")
    for i, block in enumerate(blocks):
        print(f"{i+1}. {block.content_type.value}: {block.text[:50]}...")
    
    stats = parser.get_statistics()
    print(f"\n统计信息: {stats}")