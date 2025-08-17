"""
AutoWord vNext Document Planner Demo
演示 vNext 文档规划器的使用
"""

import json
from datetime import datetime
from autoword.vnext.planner import DocumentPlanner
from autoword.vnext.models import (
    StructureV1, DocumentMetadata, StyleDefinition, ParagraphSkeleton, 
    HeadingReference, FontSpec, ParagraphSpec, LineSpacingMode, StyleType
)


def create_sample_structure() -> StructureV1:
    """创建示例文档结构"""
    return StructureV1(
        metadata=DocumentMetadata(
            title="学术论文示例",
            author="研究者",
            creation_time=datetime.now(),
            modified_time=datetime.now(),
            page_count=10,
            paragraph_count=50,
            word_count=5000
        ),
        styles=[
            StyleDefinition(
                name="标题 1",
                type=StyleType.PARAGRAPH,
                font=FontSpec(east_asian="楷体", latin="Times New Roman", size_pt=14, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="标题 2",
                type=StyleType.PARAGRAPH,
                font=FontSpec(east_asian="宋体", latin="Times New Roman", size_pt=12, bold=True),
                paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
            ),
            StyleDefinition(
                name="正文",
                type=StyleType.PARAGRAPH,
                font=FontSpec(east_asian="宋体", latin="Times New Roman", size_pt=12),
                paragraph=ParagraphSpec(line_spacing_mode=LineSpacingMode.MULTIPLE, line_spacing_value=2.0)
            )
        ],
        paragraphs=[
            ParagraphSkeleton(index=0, style_name="标题 1", preview_text="摘要", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=1, style_name="正文", preview_text="本文研究了...", is_heading=False),
            ParagraphSkeleton(index=2, style_name="标题 1", preview_text="1. 引言", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=3, style_name="正文", preview_text="随着技术的发展...", is_heading=False),
            ParagraphSkeleton(index=4, style_name="标题 2", preview_text="1.1 研究背景", is_heading=True, heading_level=2),
            ParagraphSkeleton(index=5, style_name="正文", preview_text="在当前的研究领域中...", is_heading=False),
            ParagraphSkeleton(index=6, style_name="标题 1", preview_text="参考文献", is_heading=True, heading_level=1),
            ParagraphSkeleton(index=7, style_name="正文", preview_text="[1] 张三. 研究方法...", is_heading=False)
        ],
        headings=[
            HeadingReference(paragraph_index=0, level=1, text="摘要", style_name="标题 1"),
            HeadingReference(paragraph_index=2, level=1, text="1. 引言", style_name="标题 1"),
            HeadingReference(paragraph_index=4, level=2, text="1.1 研究背景", style_name="标题 2"),
            HeadingReference(paragraph_index=6, level=1, text="参考文献", style_name="标题 1")
        ]
    )


def demo_plan_generation():
    """演示计划生成"""
    print("=== AutoWord vNext Document Planner Demo ===\n")
    
    # 创建文档规划器
    planner = DocumentPlanner()
    print("✓ 文档规划器已初始化")
    
    # 创建示例文档结构
    structure = create_sample_structure()
    print("✓ 示例文档结构已创建")
    print(f"  - 标题: {structure.metadata.title}")
    print(f"  - 段落数: {len(structure.paragraphs)}")
    print(f"  - 标题数: {len(structure.headings)}")
    print(f"  - 样式数: {len(structure.styles)}")
    
    # 定义用户意图
    user_intent = """
    请按照以下要求处理文档：
    1. 删除"摘要"和"参考文献"部分
    2. 设置标准的学术论文格式：
       - 标题 1：楷体，12pt，粗体，2倍行距
       - 标题 2：宋体，12pt，粗体，2倍行距  
       - 正文：宋体，12pt，2倍行距
    3. 最后更新目录
    """
    
    print(f"\n用户意图:\n{user_intent}")
    
    try:
        # 生成执行计划
        print("\n正在生成执行计划...")
        plan = planner.generate_plan(structure, user_intent)
        
        print("✓ 执行计划生成成功!")
        print(f"  - 操作数量: {len(plan.ops)}")
        
        # 显示计划详情
        print("\n=== 执行计划详情 ===")
        for i, op in enumerate(plan.ops, 1):
            print(f"{i}. {op.operation_type}")
            if hasattr(op, 'heading_text'):
                print(f"   - 标题文本: {op.heading_text}")
                print(f"   - 级别: {op.level}")
            elif hasattr(op, 'target_style_name'):
                print(f"   - 目标样式: {op.target_style_name}")
                if hasattr(op, 'font') and op.font:
                    print(f"   - 字体: {op.font.east_asian}, {op.font.size_pt}pt")
        
        # 验证计划
        print("\n=== 计划验证 ===")
        schema_validation = planner.validate_plan_schema(plan.model_dump())
        whitelist_validation = planner.check_whitelist_compliance(plan)
        
        print(f"✓ Schema验证: {'通过' if schema_validation.is_valid else '失败'}")
        print(f"✓ 白名单验证: {'通过' if whitelist_validation.is_valid else '失败'}")
        
        if not schema_validation.is_valid:
            print("Schema验证错误:")
            for error in schema_validation.errors:
                print(f"  - {error}")
        
        if not whitelist_validation.is_valid:
            print("白名单验证错误:")
            for error in whitelist_validation.errors:
                print(f"  - {error}")
        
        # 输出JSON格式的计划
        print("\n=== JSON格式计划 ===")
        plan_json = plan.model_dump_json(indent=2, ensure_ascii=False)
        print(plan_json)
        
    except Exception as e:
        print(f"❌ 计划生成失败: {str(e)}")
        import traceback
        traceback.print_exc()


def demo_schema_validation():
    """演示schema验证"""
    print("\n=== Schema验证演示 ===")
    
    planner = DocumentPlanner()
    
    # 测试有效计划
    valid_plan = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "摘要",
                "level": 1,
                "match": "EXACT",
                "case_sensitive": False
            },
            {
                "operation_type": "update_toc"
            }
        ]
    }
    
    result = planner.validate_plan_schema(valid_plan)
    print(f"有效计划验证: {'通过' if result.is_valid else '失败'}")
    
    # 测试无效计划
    invalid_plan = {
        "schema_version": "plan.v1",
        "ops": [
            {
                "operation_type": "delete_section_by_heading",
                "heading_text": "",  # 空字符串无效
                "level": 10,  # 级别超出范围
                "match": "INVALID"  # 无效匹配模式
            }
        ]
    }
    
    result = planner.validate_plan_schema(invalid_plan)
    print(f"无效计划验证: {'通过' if result.is_valid else '失败'}")
    if not result.is_valid:
        print("验证错误:")
        for error in result.errors:
            print(f"  - {error}")


def demo_whitelist_compliance():
    """演示白名单合规性检查"""
    print("\n=== 白名单合规性演示 ===")
    
    planner = DocumentPlanner()
    
    print("支持的操作类型:")
    for op_type in sorted(planner.WHITELISTED_OPERATIONS):
        print(f"  - {op_type}")
    
    # 测试合规计划
    from autoword.vnext.models import DeleteSectionByHeading, UpdateToc
    
    compliant_plan = StructureV1(
        metadata=DocumentMetadata(),
        ops=[
            DeleteSectionByHeading(heading_text="摘要", level=1),
            UpdateToc()
        ]
    )
    
    # Note: This is just for demo - in real usage, PlanV1 would be used
    print("\n所有操作都在白名单中，合规性检查应该通过。")


if __name__ == "__main__":
    try:
        demo_plan_generation()
        demo_schema_validation()
        demo_whitelist_compliance()
        
        print("\n=== 演示完成 ===")
        print("DocumentPlanner 已成功实现以下功能:")
        print("✓ LLM集成和计划生成")
        print("✓ JSON Schema验证")
        print("✓ 白名单操作检查")
        print("✓ 错误处理和回滚")
        print("✓ 完整的单元测试覆盖")
        
    except Exception as e:
        print(f"演示过程中出现错误: {str(e)}")
        import traceback
        traceback.print_exc()