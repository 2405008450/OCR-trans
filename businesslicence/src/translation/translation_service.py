"""Translation service for image translation system.

This module provides the TranslationService class for translating text
using the DeepSeek API with retry mechanism and batch processing support.
"""

import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Any, List, Optional

import requests

from src.config.config_manager import ConfigManager
from src.exceptions import TranslationError
from src.models.data_models import TranslationResult


logger = logging.getLogger(__name__)


class TranslationService:
    """Service for translating text using the DeepSeek API.
    
    This service provides:
    - Single text translation with automatic retry
    - Batch translation with parallel processing
    - Exponential backoff retry strategy
    - Comprehensive error handling and logging
    
    Attributes:
        config: Configuration manager instance
        api_key: DeepSeek API key
        timeout: Request timeout in seconds
        max_retries: Maximum number of retry attempts
        retry_backoff: List of backoff intervals in seconds
    """
    
    # DeepSeek API endpoint for chat completions
    API_URL = "https://api.deepseek.com/v1/chat/completions"
    
    def __init__(self, config: ConfigManager):
        """Initialize the translation service.
        
        Args:
            config: Configuration manager instance
        """
        self.config = config
        self.api_key = config.get('api.deepseek_key', '')
        self.timeout = config.get('api.timeout', 10)
        self.max_retries = config.get('api.max_retries', 3)
        self.retry_backoff = config.get('api.retry_backoff', [1, 2, 4])
        self.max_workers = config.get('performance.max_workers', 4)
        
        logger.debug(
            f"TranslationService initialized with timeout={self.timeout}, "
            f"max_retries={self.max_retries}"
        )

    def translate(
        self, 
        text: str, 
        source_lang: str = "zh", 
        target_lang: str = "en"
    ) -> TranslationResult:
        """Translate text from source language to target language.
        
        Uses the DeepSeek API to translate text with automatic retry
        on failure using exponential backoff.
        
        Args:
            text: Source text to translate
            source_lang: Source language code (default: "zh" for Chinese)
            target_lang: Target language code (default: "en" for English)
            
        Returns:
            TranslationResult containing the translation or error information
        """
        if not text or not text.strip():
            logger.warning("Empty text provided for translation")
            return TranslationResult(
                source_text=text,
                translated_text="",
                confidence=0.0,
                success=True,
                error_message=None
            )
        
        logger.info(f"Translating text: '{text[:50]}...' from {source_lang} to {target_lang}")
        
        def api_call() -> TranslationResult:
            return self._call_api(text, source_lang, target_lang)
        
        try:
            result = self._retry_with_backoff(
                api_call, 
                self.max_retries, 
                self.retry_backoff
            )
            return result
        except Exception as e:
            error_msg = f"Translation failed after {self.max_retries} retries: {str(e)}"
            logger.error(error_msg)
            return TranslationResult(
                source_text=text,
                translated_text="",
                confidence=0.0,
                success=False,
                error_message=error_msg
            )
    
    def _call_api(
        self, 
        text: str, 
        source_lang: str, 
        target_lang: str
    ) -> TranslationResult:
        """Make a single API call to translate text.
        
        Args:
            text: Source text to translate
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            TranslationResult with translation or error
            
        Raises:
            TranslationError: If the API call fails
        """
        if not self.api_key:
            raise TranslationError("DeepSeek API key is not configured")
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        # Construct the translation prompt
        lang_names = {
            "zh": "Chinese",
            "en": "English",
            "ja": "Japanese",
            "ko": "Korean",
            "fr": "French",
            "de": "German",
            "es": "Spanish"
        }
        
        source_name = lang_names.get(source_lang, source_lang)
        target_name = lang_names.get(target_lang, target_lang)
        
        # 改进的翻译提示词：提供更多上下文和指导
        prompt = (
            f"Translate the following {source_name} text to {target_name}.\n\n"
            f"Requirements:\n"
            f"1. Provide ONLY the translation, no explanations or notes\n"
            f"2. Preserve all proper nouns (names, places, organizations)\n"
            f"3. Keep all numbers, dates, and codes exactly as they are\n"
            f"4. Maintain professional and formal tone for official documents\n"
            f"5. Translate legal/business terms accurately (e.g., '有限责任公司' = 'Limited Liability Company')\n"
            f"6. If the text is incomplete or unclear, translate what you can see\n\n"
            f"Text to translate:\n{text}"
        )
        
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "You are a professional translator specializing in official documents, "
                        "legal texts, and business materials. Translate accurately, maintaining "
                        "formal tone and proper terminology. Always preserve proper nouns, numbers, "
                        "and formatting exactly as they appear."
                    )
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.1,  # 降低温度以获得更一致、准确的翻译
            "max_tokens": 2048
        }
        
        try:
            response = requests.post(
                self.API_URL,
                headers=headers,
                json=payload,
                timeout=self.timeout
            )
            
            response.raise_for_status()
            
            result = response.json()
            translated_text = self._parse_response(result)
            
            logger.debug(f"Translation successful: '{translated_text[:50]}...'")
            
            return TranslationResult(
                source_text=text,
                translated_text=translated_text,
                confidence=0.95,  # DeepSeek doesn't provide confidence, use default
                success=True,
                error_message=None
            )
            
        except requests.exceptions.Timeout:
            raise TranslationError(f"API request timed out after {self.timeout} seconds")
        except requests.exceptions.HTTPError as e:
            status_code = e.response.status_code if e.response else "unknown"
            error_body = ""
            try:
                error_body = e.response.json() if e.response else {}
            except Exception:
                pass
            raise TranslationError(
                f"API returned error status {status_code}: {error_body}"
            )
        except requests.exceptions.RequestException as e:
            raise TranslationError(f"API request failed: {str(e)}")
        except Exception as e:
            raise TranslationError(f"Unexpected error during translation: {str(e)}")
    
    def _parse_response(self, response: dict) -> str:
        """Parse the API response to extract translated text.
        
        Args:
            response: JSON response from the API
            
        Returns:
            Translated text string
            
        Raises:
            TranslationError: If response format is invalid
        """
        try:
            choices = response.get("choices", [])
            if not choices:
                raise TranslationError("API response contains no choices")
            
            message = choices[0].get("message", {})
            content = message.get("content", "")
            
            if not content:
                raise TranslationError("API response contains empty content")
            
            return content.strip()
            
        except KeyError as e:
            raise TranslationError(f"Invalid API response format: missing key {e}")

    def _retry_with_backoff(
        self, 
        func: Callable[[], Any], 
        max_retries: int, 
        backoff: List[int]
    ) -> Any:
        """Execute a function with exponential backoff retry strategy.
        
        Retries the function on failure with increasing delays between attempts.
        Logs detailed information about each retry attempt.
        
        Args:
            func: Function to execute (should take no arguments)
            max_retries: Maximum number of retry attempts
            backoff: List of backoff intervals in seconds
            
        Returns:
            Result of the function call
            
        Raises:
            The last exception if all retries fail
        """
        last_exception = None
        
        for attempt in range(max_retries + 1):
            try:
                result = func()
                if attempt > 0:
                    logger.info(f"Retry attempt {attempt} succeeded")
                return result
                
            except Exception as e:
                last_exception = e
                
                if attempt < max_retries:
                    # Get backoff time for this attempt
                    backoff_time = backoff[attempt] if attempt < len(backoff) else backoff[-1]
                    
                    logger.warning(
                        f"Attempt {attempt + 1}/{max_retries + 1} failed: {str(e)}. "
                        f"Retrying in {backoff_time} seconds..."
                    )
                    
                    time.sleep(backoff_time)
                else:
                    logger.error(
                        f"All {max_retries + 1} attempts failed. Last error: {str(e)}"
                    )
        
        # All retries exhausted, raise the last exception
        raise last_exception
    
    def translate_batch(
        self, 
        texts: List[str], 
        source_lang: str = "zh", 
        target_lang: str = "en"
    ) -> List[TranslationResult]:
        """Translate multiple texts in parallel.
        
        Uses ThreadPoolExecutor to process multiple translations concurrently.
        Each text is translated independently, so failures in one don't affect others.
        
        Args:
            texts: List of source texts to translate
            source_lang: Source language code (default: "zh")
            target_lang: Target language code (default: "en")
            
        Returns:
            List of TranslationResult objects in the same order as input texts
        """
        if not texts:
            logger.warning("Empty text list provided for batch translation")
            return []
        
        logger.info(f"Starting batch translation of {len(texts)} texts")
        
        results: List[Optional[TranslationResult]] = [None] * len(texts)
        
        # Check if parallel translation is enabled
        parallel_enabled = self.config.get('performance.parallel_translation', True)
        
        if parallel_enabled and len(texts) > 1:
            # Use thread pool for parallel processing
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all translation tasks
                future_to_index = {
                    executor.submit(
                        self.translate, text, source_lang, target_lang
                    ): i
                    for i, text in enumerate(texts)
                }
                
                # Collect results as they complete
                for future in as_completed(future_to_index):
                    index = future_to_index[future]
                    try:
                        results[index] = future.result()
                    except Exception as e:
                        logger.error(f"Batch translation failed for text {index}: {str(e)}")
                        results[index] = TranslationResult(
                            source_text=texts[index],
                            translated_text="",
                            confidence=0.0,
                            success=False,
                            error_message=str(e)
                        )
        else:
            # Sequential processing
            for i, text in enumerate(texts):
                try:
                    results[i] = self.translate(text, source_lang, target_lang)
                except Exception as e:
                    logger.error(f"Batch translation failed for text {i}: {str(e)}")
                    results[i] = TranslationResult(
                        source_text=text,
                        translated_text="",
                        confidence=0.0,
                        success=False,
                        error_message=str(e)
                    )
        
        # Count successes and failures
        successes = sum(1 for r in results if r and r.success)
        failures = len(texts) - successes
        
        logger.info(
            f"Batch translation completed: {successes} succeeded, {failures} failed"
        )
        
        return results
