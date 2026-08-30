"""
LLM Handler — stub for Community Edition.

The enterprise Email Report and Data Chat features are not included in this
edition.  This module exports the same public names so that any import line
that was not updated yet does not raise an ImportError, but every method
raises NotImplementedError at call-time.
"""
from typing import Dict, Any, Optional, Callable, List
import pandas as pd


class LLMHandler:
    """
    Stub — enterprise feature not available in Community Edition.
    """

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.llm = None
        self.llm_model_name = None
        self.loaded_gpu_enabled = None

    @property
    def is_model_available(self) -> bool:
        return False

    def load_model(self, force: bool = False) -> bool:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def preload_model_async(self, callback: Optional[Callable[[bool, str], None]] = None) -> None:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def unload_model(self) -> None:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def build_data_context(self, df: pd.DataFrame, cache: Optional[Dict[str, Any]] = None) -> str:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def chat(self, question: str, df: pd.DataFrame,
             cache: Optional[Dict[str, Any]] = None) -> str:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def chat_async(self, question: str, df: pd.DataFrame,
                   callback: Callable[[str, Optional[str]], None],
                   cache: Optional[Dict[str, Any]] = None) -> None:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def generate_email(self, context: str) -> str:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def generate_email_async(self, context: str,
                             callback: Callable[[str], None]) -> None:
        raise NotImplementedError("LLM features are not available in Community Edition.")

    def build_email_context(self, df: pd.DataFrame, cache: Dict[str, Any],
                            column_keywords: Dict[str, List[str]],
                            null_threshold: float = 20.0) -> str:
        raise NotImplementedError("LLM features are not available in Community Edition.")
