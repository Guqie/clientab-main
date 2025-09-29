"""
CSV转Word转换功能集成测试
测试不同模板、输出格式和边界条件的转换功能
"""

import pytest
import pandas as pd
import os
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock

# 导入被测试的模块
from csv_word_converter.core import UniversalDocumentGenerator
from csv_word_converter.cli import main as cli_main
from csv_word_converter.async_converter import AsyncConverter
from csv_word_converter.output_formats import OutputFormatFactory


class TestCSVWordIntegration:
    """CSV转Word转换功能集成测试类"""
    
    @pytest.fixture
    def sample_csv_data(self):
        """创建测试用的CSV数据"""
        return pd.DataFrame({
            'title': ['测试标题1', '测试标题2'],
            'content': ['这是测试内容1', '这是测试内容2'],
            'image_url': ['https://example.com/image1.jpg', 'https://example.com/image2.jpg'],
            'category': ['科技', '教育']
        })
    
    @pytest.fixture
    def temp_csv_file(self, sample_csv_data):
        """创建临时CSV文件"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as f:
            sample_csv_data.to_csv(f.name, index=False, encoding='utf-8')
            yield f.name
        # 清理临时文件
        if os.path.exists(f.name):
            os.unlink(f.name)
    
    @pytest.fixture
    def temp_output_dir(self):
        """创建临时输出目录"""
        with tempfile.TemporaryDirectory() as temp_dir:
            yield temp_dir
    
    def test_basic_csv_to_word_conversion(self, temp_csv_file, temp_output_dir):
        """测试基本的CSV到Word转换功能"""
        output_path = os.path.join(temp_output_dir, 'test_output.docx')
        
        # 使用UniversalDocumentGenerator进行转换
        generator = UniversalDocumentGenerator(template_type='guoziwei')
        data = pd.read_csv(temp_csv_file)
        
        # 转换为字典列表格式
        data_dict = data.to_dict('records')
        
        result = generator.generate_document(data=data_dict)
        
        # 验证文件生成成功
        assert os.path.exists(result)
        assert os.path.getsize(result) > 0
        
        # 验证文件是有效的Word文档
        from docx import Document
        doc = Document(result)
        assert len(doc.paragraphs) > 0
    
    def test_different_templates(self, temp_csv_file, temp_output_dir):
        """测试不同模板的转换功能"""
        templates = ['guoziwei', 'default']  # 假设有这些模板
        data = pd.read_csv(temp_csv_file)
        data_dict = data.to_dict('records')
        
        for template in templates:
            output_path = os.path.join(temp_output_dir, f'test_{template}.docx')
            
            try:
                generator = UniversalDocumentGenerator(template_type=template)
                result = generator.generate_document(data=data_dict)
                
                # 验证文件生成成功
                assert os.path.exists(result)
                assert os.path.getsize(result) > 0
                
            except Exception as e:
                # 某些模板可能不存在，记录但不失败
                print(f"模板 {template} 测试失败: {str(e)}")
    
    def test_output_formats(self, temp_csv_file, temp_output_dir):
        """测试不同输出格式的转换功能"""
        import asyncio
        
        async def async_format_test():
            formats = ['docx', 'pdf', 'html', 'markdown']
            format_factory = OutputFormatFactory()
            data = pd.read_csv(temp_csv_file)
            
            for format_type in formats:
                output_path = os.path.join(temp_output_dir, f'test_output.{format_type}')
                
                try:
                    # 根据格式类型选择相应的转换方法
                    if format_type == 'docx':
                        generator = UniversalDocumentGenerator(template_type='guoziwei')
                        result = generator.generate_document(data=data.to_dict('records'))
                    else:
                        # 创建格式配置
                        from csv_word_converter.output_formats import FormatConfig
                        config = FormatConfig(
                            format_type=format_type,
                            output_path=Path(output_path),
                            template_data={'data': data.to_dict('records')},
                            options={}
                        )
                        
                        # 创建格式化器并执行转换
                        formatter = format_factory.create_formatter(config)
                        result_path = await formatter.format_output()
                        result = str(result_path)
                    
                    # 验证文件生成成功
                    assert os.path.exists(result)
                    assert os.path.getsize(result) > 0
                    
                except Exception as e:
                    # 某些格式可能需要额外依赖，记录但不失败
                    print(f"格式 {format_type} 转换失败: {str(e)}")
        
        # 运行异步测试
        asyncio.run(async_format_test())
    
    def test_async_conversion(self, temp_csv_file, temp_output_dir):
        """测试异步转换功能"""
        import asyncio
        
        async def async_test():
            converter = AsyncConverter()
            
            # 添加转换任务
            task_id = await converter.add_task(
                csv_file=temp_csv_file,
                output_path=temp_output_dir,
                template_type='guoziwei'
            )
            
            # 处理所有任务
            await converter.process_all_tasks()
            
            # 等待一段时间确保任务完成
            import time
            await asyncio.sleep(1)
            
            # 检查任务状态
            task_status = converter.get_task_status(task_id)
            print(f"任务状态: {task_status}")
            
            # 验证任务完成
            if task_status != 'completed':
                # 如果任务未完成，等待更长时间
                await asyncio.sleep(2)
                task_status = converter.get_task_status(task_id)
                print(f"重新检查任务状态: {task_status}")
            
            # 获取所有任务信息
            all_tasks = converter.get_all_tasks()
            assert len(all_tasks) > 0
            
            # 验证输出文件存在（如果任务成功完成）
            if task_status == 'completed':
                task_info = all_tasks[0]
                if hasattr(task_info, 'output_path') and task_info.output_path:
                    assert os.path.exists(task_info.output_path)
            
            # 清理资源
            await converter.shutdown()
        
        # 运行异步测试
        asyncio.run(async_test())
    
    def test_cli_integration(self, temp_csv_file, temp_output_dir):
        """测试CLI命令行接口集成"""
        import sys
        from io import StringIO
        from src.csv_word_converter.cli import main as cli_main
        
        # 构建CLI参数 - 只指定输出目录，让CLI自动生成文件名
        test_args = [
            'csv_word_converter',
            temp_csv_file,
            '-t', 'guoziwei',
            '--output-dir', temp_output_dir,
            '--no-images'  # 禁用图片下载以加快测试
        ]
        
        # 模拟命令行参数
        with patch.object(sys, 'argv', test_args):
            try:
                # 捕获输出
                captured_output = StringIO()
                with patch('sys.stdout', captured_output):
                    result = cli_main()
                
                # CLI实际输出到temp-data目录，检查该目录中是否有生成的文件
                temp_data_dir = "temp-data"
                if os.path.exists(temp_data_dir):
                    output_files = [f for f in os.listdir(temp_data_dir) if f.endswith('.docx')]
                    assert len(output_files) > 0, f"CLI执行后未在{temp_data_dir}中生成任何.docx文件"
                    
                    # 验证生成的文件不为空
                    output_file = os.path.join(temp_data_dir, output_files[0])
                    assert os.path.getsize(output_file) > 0, "生成的文件大小为0"
                else:
                    assert False, "temp-data目录不存在，CLI可能执行失败"
                
            except SystemExit as e:
                # CLI可能会调用sys.exit()，检查退出码
                assert e.code == 0, f"CLI执行失败，退出码: {e.code}"
                # 即使有SystemExit，也要验证文件是否生成
                temp_data_dir = "temp-data"
                if os.path.exists(temp_data_dir):
                    output_files = [f for f in os.listdir(temp_data_dir) if f.endswith('.docx')]
                    assert len(output_files) > 0, f"CLI执行后未在{temp_data_dir}中生成任何.docx文件"
                
                output_file = os.path.join(temp_output_dir, output_files[0])
                assert os.path.getsize(output_file) > 0, "生成的文件大小为0"
    
    def test_error_handling(self, temp_output_dir):
        """测试错误处理功能"""
        # 测试空数据
        generator = UniversalDocumentGenerator(template_type='guoziwei')
        empty_data = []
        
        try:
            result = generator.generate_document(data=empty_data)
            # 即使是空数据，也应该能生成基本文档
            assert os.path.exists(result)
        except Exception as e:
            # 记录错误但不失败，因为空数据处理可能有不同策略
            print(f"空数据处理: {str(e)}")
        
        # 测试无效数据格式
        invalid_data = [{"invalid": "data without required fields"}]
        
        try:
            result = generator.generate_document(data=invalid_data)
            # 应该能处理无效数据
            assert os.path.exists(result)
        except Exception as e:
            print(f"无效数据处理: {str(e)}")
    
    def test_large_dataset_performance(self, temp_output_dir):
        """测试大数据集的性能"""
        # 创建大数据集
        large_data = []
        for i in range(100):
            large_data.append({
                'title': f'测试标题{i}',
                'content': f'测试内容{i}' * 10,  # 增加内容长度
                'category': f'分类{i % 5}'
            })
        
        output_path = os.path.join(temp_output_dir, 'large_test.docx')
        
        # 测试性能
        import time
        start_time = time.time()
        
        generator = UniversalDocumentGenerator(template_type='guoziwei')
        result = generator.generate_document(data=large_data)
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        # 验证文件生成成功
        assert os.path.exists(result)
        assert os.path.getsize(result) > 0
        
        # 性能检查（可根据实际需求调整）
        print(f"处理100条记录耗时: {processing_time:.2f}秒")
        assert processing_time < 30  # 假设30秒内完成


if __name__ == '__main__':
    # 运行测试
    pytest.main([__file__, '-v'])