"""
統一日誌配置和錯誤處理模組
提供一致的日誌記錄和錯誤處理機制
"""

import logging
import sys
from typing import Optional, Any, Dict
from pathlib import Path
from datetime import datetime
from enum import Enum


class LogLevel(Enum):
    """日誌級別枚舉"""
    DEBUG = logging.DEBUG
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    CRITICAL = logging.CRITICAL


class ConversionError(Exception):
    """轉換相關錯誤的基類"""
    def __init__(self, message: str, error_code: Optional[str] = None, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.error_code = error_code
        self.details = details or {}
        self.timestamp = datetime.now()


class DocumentError(ConversionError):
    """文件處理錯誤"""
    pass


class FormatError(ConversionError):
    """格式處理錯誤"""
    pass


class SlideError(ConversionError):
    """投影片操作錯誤"""
    pass


class LoggerConfig:
    """日誌配置管理器"""
    
    @staticmethod
    def setup_logger(
        name: str = "document_converter",
        level: LogLevel = LogLevel.INFO,
        log_file: Optional[str] = None,
        console_output: bool = True,
        file_output: bool = True
    ) -> logging.Logger:
        """
        設置日誌記錄器
        
        Args:
            name: 日誌記錄器名稱
            level: 日誌級別
            log_file: 日誌檔案路徑（None 則使用預設）
            console_output: 是否輸出到控制台
            file_output: 是否輸出到檔案
            
        Returns:
            logging.Logger: 配置好的日誌記錄器
        """
        logger = logging.getLogger(name)
        logger.setLevel(level.value)
        
        # 清除現有的處理器
        logger.handlers.clear()
        
        # 設置格式器
        formatter = LoggerConfig._get_formatter()
        
        # 控制台處理器
        if console_output:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(level.value)
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        # 檔案處理器
        if file_output:
            if log_file is None:
                log_file = LoggerConfig._get_default_log_file()
            
            # 確保日誌目錄存在
            log_path = Path(log_file)
            log_path.parent.mkdir(parents=True, exist_ok=True)
            
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(level.value)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        # 避免重複記錄
        logger.propagate = False
        
        return logger
    
    @staticmethod
    def _get_formatter() -> logging.Formatter:
        """獲取日誌格式器"""
        return logging.Formatter(
            fmt='%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
    
    @staticmethod
    def _get_default_log_file() -> str:
        """獲取預設日誌檔案路徑"""
        timestamp = datetime.now().strftime("%Y%m%d")
        return f"logs/document_converter_{timestamp}.log"


class ErrorHandler:
    """統一錯誤處理器"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """
        初始化錯誤處理器
        
        Args:
            logger: 日誌記錄器
        """
        self.logger = logger or LoggerConfig.setup_logger()
    
    def handle_error(self, error: Exception, context: str = "", 
                    error_code: Optional[str] = None) -> Dict[str, Any]:
        """
        處理錯誤並記錄
        
        Args:
            error: 錯誤物件
            context: 錯誤上下文
            error_code: 錯誤代碼
            
        Returns:
            Dict: 錯誤資訊字典
        """
        error_info = {
            'error_type': type(error).__name__,
            'error_message': str(error),
            'error_code': error_code,
            'context': context,
            'timestamp': datetime.now().isoformat()
        }
        
        # 如果是自定義錯誤，添加額外資訊
        if isinstance(error, ConversionError):
            error_info.update({
                'error_code': error.error_code,
                'details': error.details
            })
        
        # 記錄錯誤
        if isinstance(error, (DocumentError, FormatError, SlideError)):
            self.logger.error(f"[{context}] {error_info['error_type']}: {error_info['error_message']}")
        else:
            self.logger.exception(f"[{context}] 未預期錯誤: {error_info['error_message']}")
        
        return error_info
    
    def log_operation_start(self, operation: str, details: Optional[Dict[str, Any]] = None):
        """記錄操作開始"""
        message = f"開始操作: {operation}"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" ({detail_str})"
        self.logger.info(message)
    
    def log_operation_success(self, operation: str, result: Optional[Dict[str, Any]] = None):
        """記錄操作成功"""
        message = f"操作成功: {operation}"
        if result:
            result_str = ", ".join([f"{k}={v}" for k, v in result.items() if k not in ['success', 'error']])
            if result_str:
                message += f" ({result_str})"
        self.logger.info(message)
    
    def log_operation_warning(self, operation: str, warning: str, details: Optional[Dict[str, Any]] = None):
        """記錄操作警告"""
        message = f"操作警告: {operation} - {warning}"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" ({detail_str})"
        self.logger.warning(message)
    
    def log_progress(self, current: int, total: int, operation: str = ""):
        """記錄進度"""
        percentage = (current / total) * 100 if total > 0 else 0
        message = f"進度: {current}/{total} ({percentage:.1f}%)"
        if operation:
            message += f" - {operation}"
        self.logger.debug(message)


class PerformanceMonitor:
    """性能監控器"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """
        初始化性能監控器
        
        Args:
            logger: 日誌記錄器
        """
        self.logger = logger or LoggerConfig.setup_logger()
        self.start_times = {}
    
    def start_timing(self, operation: str):
        """開始計時"""
        self.start_times[operation] = datetime.now()
        self.logger.debug(f"開始計時: {operation}")
    
    def end_timing(self, operation: str) -> float:
        """結束計時並返回耗時（秒）"""
        if operation not in self.start_times:
            self.logger.warning(f"未找到操作的開始時間: {operation}")
            return 0.0
        
        start_time = self.start_times.pop(operation)
        duration = (datetime.now() - start_time).total_seconds()
        
        self.logger.info(f"操作完成: {operation} (耗時: {duration:.2f}秒)")
        return duration
    
    def log_memory_usage(self, operation: str = ""):
        """記錄記憶體使用情況（如果可用）"""
        try:
            import psutil
            process = psutil.Process()
            memory_info = process.memory_info()
            memory_mb = memory_info.rss / 1024 / 1024
            
            message = f"記憶體使用: {memory_mb:.1f}MB"
            if operation:
                message += f" - {operation}"
            
            self.logger.debug(message)
            
        except ImportError:
            # psutil 不可用，跳過記憶體監控
            pass
        except Exception as e:
            self.logger.warning(f"無法獲取記憶體使用情況: {e}")


def create_result_dict(success: bool = True, error: Optional[str] = None, 
                      **kwargs) -> Dict[str, Any]:
    """
    創建標準結果字典
    
    Args:
        success: 操作是否成功
        error: 錯誤訊息
        **kwargs: 其他結果數據
        
    Returns:
        Dict: 標準格式的結果字典
    """
    result = {
        'success': success,
        'error': error,
        'timestamp': datetime.now().isoformat()
    }
    result.update(kwargs)
    return result


# 全域日誌記錄器實例
default_logger = LoggerConfig.setup_logger()
default_error_handler = ErrorHandler(default_logger)
default_performance_monitor = PerformanceMonitor(default_logger)


def get_logger(name: str = "document_converter") -> logging.Logger:
    """
    獲取日誌記錄器（便利函數）
    
    Args:
        name: 日誌記錄器名稱
        
    Returns:
        logging.Logger: 日誌記錄器
    """
    return LoggerConfig.setup_logger(name)