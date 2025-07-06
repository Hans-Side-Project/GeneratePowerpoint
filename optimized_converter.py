"""
優化版本的轉換器 - 使用新的模組化架構
這個檔案替代原來的 word_to_ppt_converter.py，提供更好的性能和錯誤處理
"""

import sys
import os
from typing import Optional, Callable

# 確保可以導入我們的模組
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from document_converter import convert_word_to_ppt, analyze_document_structure, ConverterFactory
from logger_config import get_logger, LogLevel


def main():
    """
    主函數：執行 Word 到 PowerPoint 的轉換
    """
    # 設置日誌
    logger = get_logger("OptimizedConverter")
    
    # 檔案路徑
    word_file = "證道資料.docx"
    ppt_file = "證道資料.pptx"
    
    print("🔄 開始使用優化版轉換器...")
    logger.info("啟動優化版轉換器")
    
    # 檢查檔案是否存在
    if not os.path.exists(word_file):
        print(f"❌ Word 檔案不存在: {word_file}")
        logger.error(f"Word 檔案不存在: {word_file}")
        return
    
    if not os.path.exists(ppt_file):
        print(f"❌ PowerPoint 模板不存在: {ppt_file}")
        logger.error(f"PowerPoint 模板不存在: {ppt_file}")
        return
    
    # 進度回調函數
    def progress_callback(current: float, total: float, message: str):
        percentage = (current / total) * 100 if total > 0 else 0
        print(f"📊 進度: {percentage:.1f}% - {message}")
    
    try:
        # 1. 先分析文檔結構
        print("\n🔍 分析文檔結構...")
        word_analysis = analyze_document_structure(word_file)
        
        if word_analysis['success']:
            print(f"✅ Word 文檔分析完成:")
            print(f"   📄 總段落數: {word_analysis['total_sections']}")
            print(f"   📝 總字符數: {len(word_analysis['basic_content']['text'])}")
            
            # 顯示前幾個段落預覽
            sections = word_analysis['sections'][:3]
            for section in sections:
                title = section['title'][:30] + ('...' if len(section['title']) > 30 else '')
                print(f"   {section['number']}. {title}")
        else:
            print(f"❌ Word 文檔分析失敗: {word_analysis['error']}")
            return
        
        # 2. 分析模板
        print(f"\n🔍 分析 PowerPoint 模板...")
        ppt_analysis = analyze_document_structure(ppt_file)
        
        if ppt_analysis['success']:
            print(f"✅ PowerPoint 模板分析完成:")
            print(f"   📊 模板投影片數: {ppt_analysis['total_slides']}")
            if 'structure_analysis' in ppt_analysis:
                structure = ppt_analysis['structure_analysis']
                print(f"   🎨 使用的版面: {', '.join(structure['layouts_used'])}")
                print(f"   📝 文本框總數: {structure['text_shapes_count']}")
        else:
            print(f"❌ PowerPoint 模板分析失敗: {ppt_analysis['error']}")
            return
        
        # 3. 執行轉換
        print(f"\n🚀 開始轉換...")
        result = convert_word_to_ppt(word_file, ppt_file, progress_callback=progress_callback)
        
        # 4. 顯示結果
        print(f"\n" + "="*60)
        if result['success']:
            print(f"✅ 轉換成功!")
            print(f"📊 處理統計:")
            print(f"   📄 處理段落數: {result['total_sections']}")
            print(f"   📈 創建投影片數: {result['slides_created']}")
            print(f"   ⏱️  處理時間: {result.get('processing_time', 0):.2f} 秒")
            print(f"   💾 輸出檔案: {result['output_file']}")
            
            # 顯示警告信息
            if result.get('skipped_sections'):
                print(f"\n⚠️  跳過的段落:")
                for skipped in result['skipped_sections']:
                    print(f"   段落 {skipped['number']}: {skipped.get('title', '')[:30]}... ({skipped['error']})")
            
            if result.get('format_issues'):
                print(f"\n⚠️  格式問題:")
                for issue in result['format_issues']:
                    print(f"   - {issue}")
            
            print(f"\n🎯 請檢查輸出檔案: {result['output_file']}")
            
        else:
            print(f"❌ 轉換失敗!")
            print(f"錯誤信息: {result['error']}")
            
            if 'error_info' in result:
                error_info = result['error_info']
                print(f"錯誤類型: {error_info.get('error_type', 'Unknown')}")
                print(f"錯誤代碼: {error_info.get('error_code', 'N/A')}")
                
                if 'details' in error_info:
                    print(f"詳細信息: {error_info['details']}")
        
        print("="*60)
        
    except Exception as e:
        print(f"❌ 執行過程中發生未預期錯誤: {str(e)}")
        logger.exception("執行過程中發生未預期錯誤")


def demo_batch_conversion():
    """
    演示批次轉換功能
    """
    print("\n🔄 演示批次轉換功能...")
    
    # 創建批次轉換器
    batch_converter = ConverterFactory.create_batch_converter(logger_level="INFO")
    
    # 準備轉換列表（這裡只是示例）
    file_pairs = [
        {
            'source': '證道資料.docx',
            'template': '證道資料.pptx',
            'output': '批次輸出1.pptx'
        }
        # 可以添加更多檔案對
    ]
    
    def batch_progress(current: int, total: int, message: str):
        print(f"📊 批次進度: {current}/{total} - {message}")
    
    # 執行批次轉換
    results = batch_converter.convert_multiple(file_pairs, batch_progress)
    
    # 顯示批次結果
    successful = sum(1 for r in results if r['success'])
    print(f"\n📊 批次轉換完成: {successful}/{len(results)} 成功")
    
    for i, result in enumerate(results):
        if result['success']:
            print(f"✅ 檔案 {i+1}: 成功 - {result.get('output_file', 'N/A')}")
        else:
            print(f"❌ 檔案 {i+1}: 失敗 - {result['error']}")


def demo_advanced_features():
    """
    演示進階功能
    """
    print("\n🔧 演示進階功能...")
    
    # 創建自定義轉換器
    converter = ConverterFactory.create_converter(logger_level="DEBUG", log_to_file=True)
    
    # 獲取轉換預覽
    if os.path.exists("證道資料.docx"):
        print("📋 獲取轉換預覽...")
        preview = converter.get_conversion_preview("證道資料.docx")
        
        if preview['success']:
            print(f"✅ 預覽信息:")
            print(f"   估計投影片數: {preview['estimated_slides']}")
            print(f"   段落預覽:")
            for section in preview['sections_preview']:
                print(f"     {section['number']}. {section['title']} (長度: {section['content_length']})")
        else:
            print(f"❌ 預覽失敗: {preview['error']}")


if __name__ == "__main__":
    # 執行主要轉換
    main()
    
    # 可選：演示其他功能
    if len(sys.argv) > 1 and sys.argv[1] == "--demo":
        demo_advanced_features()
        demo_batch_conversion()