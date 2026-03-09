"""
Retry strategy implementation with exponential backoff
"""

import asyncio
import logging
import random
from typing import Callable, Any, Optional, Type, Union, List
from ..core.exceptions import RetryableError, ModelError

logger = logging.getLogger(__name__)


class RetryStrategy:
    """
    Implements exponential backoff retry strategy
    
    This class provides configurable retry logic with exponential backoff
    for handling transient failures in AI model requests.
    """
    
    def __init__(
        self,
        max_retries: int = 3,
        base_delay: float = 1.0,
        max_delay: float = 60.0,
        backoff_factor: float = 2.0,
        jitter: bool = True,
        retryable_exceptions: Optional[List[Type[Exception]]] = None
    ):
        """
        Initialize retry strategy
        
        Args:
            max_retries: Maximum number of retry attempts
            base_delay: Base delay in seconds for first retry
            max_delay: Maximum delay in seconds between retries
            backoff_factor: Multiplier for exponential backoff
            jitter: Whether to add random jitter to delays
            retryable_exceptions: List of exception types that should trigger retries
        """
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.max_delay = max_delay
        self.backoff_factor = backoff_factor
        self.jitter = jitter
        self.retryable_exceptions = retryable_exceptions or [RetryableError, ModelError]
        self.logger = logging.getLogger(__name__)
    
    def calculate_delay(self, attempt: int) -> float:
        """
        Calculate delay for given attempt number
        
        Args:
            attempt: Current attempt number (0-based)
            
        Returns:
            float: Delay in seconds
        """
        # Calculate exponential backoff delay
        delay = self.base_delay * (self.backoff_factor ** attempt)
        
        # Cap at maximum delay
        delay = min(delay, self.max_delay)
        
        # Add jitter if enabled
        if self.jitter:
            # Add random jitter of ±25%
            jitter_range = delay * 0.25
            delay += random.uniform(-jitter_range, jitter_range)
            delay = max(0.1, delay)  # Ensure minimum delay
        
        return delay
    
    def is_retryable(self, exception: Exception) -> bool:
        """
        Check if an exception should trigger a retry
        
        Args:
            exception: The exception to check
            
        Returns:
            bool: True if the exception is retryable
        """
        return any(isinstance(exception, exc_type) for exc_type in self.retryable_exceptions)
    
    async def execute_with_retry(
        self,
        func: Callable,
        *args,
        **kwargs
    ) -> Any:
        """
        Execute a function with retry logic
        
        Args:
            func: The function to execute
            *args: Positional arguments for the function
            **kwargs: Keyword arguments for the function
            
        Returns:
            Any: Result of the function call
            
        Raises:
            Exception: The last exception if all retries fail
        """
        last_exception = None
        
        for attempt in range(self.max_retries + 1):  # +1 for initial attempt
            try:
                if attempt > 0:
                    self.logger.info(f"Retry attempt {attempt}/{self.max_retries}")
                
                # Execute the function
                if asyncio.iscoroutinefunction(func):
                    result = await func(*args, **kwargs)
                else:
                    result = func(*args, **kwargs)
                
                if attempt > 0:
                    self.logger.info(f"Retry attempt {attempt} succeeded")
                
                return result
                
            except Exception as e:
                last_exception = e
                
                # Check if we should retry
                if not self.is_retryable(e):
                    self.logger.info(f"Exception {type(e).__name__} is not retryable, giving up")
                    raise e
                
                # Check if we have more attempts
                if attempt >= self.max_retries:
                    self.logger.error(f"All {self.max_retries + 1} attempts failed")
                    break
                
                # Calculate delay and wait
                delay = self.calculate_delay(attempt)
                self.logger.warning(
                    f"Attempt {attempt + 1} failed with {type(e).__name__}: {str(e)}. "
                    f"Retrying in {delay:.2f} seconds..."
                )
                
                await asyncio.sleep(delay)
        
        # If we get here, all retries failed
        self.logger.error(f"Function failed after {self.max_retries + 1} attempts")
        raise last_exception
    
    def get_config(self) -> dict:
        """
        Get current retry configuration
        
        Returns:
            dict: Current configuration
        """
        return {
            "max_retries": self.max_retries,
            "base_delay": self.base_delay,
            "max_delay": self.max_delay,
            "backoff_factor": self.backoff_factor,
            "jitter": self.jitter,
            "retryable_exceptions": [exc.__name__ for exc in self.retryable_exceptions]
        }
    
    def update_config(
        self,
        max_retries: Optional[int] = None,
        base_delay: Optional[float] = None,
        max_delay: Optional[float] = None,
        backoff_factor: Optional[float] = None,
        jitter: Optional[bool] = None
    ):
        """
        Update retry configuration
        
        Args:
            max_retries: New maximum retry count
            base_delay: New base delay
            max_delay: New maximum delay
            backoff_factor: New backoff factor
            jitter: New jitter setting
        """
        if max_retries is not None:
            self.max_retries = max_retries
        if base_delay is not None:
            self.base_delay = base_delay
        if max_delay is not None:
            self.max_delay = max_delay
        if backoff_factor is not None:
            self.backoff_factor = backoff_factor
        if jitter is not None:
            self.jitter = jitter
        
        self.logger.info(f"Updated retry configuration: {self.get_config()}")


class CircuitBreaker:
    """
    Circuit breaker pattern implementation
    
    Prevents cascading failures by temporarily stopping requests to a failing service.
    """
    
    def __init__(
        self,
        failure_threshold: int = 5,
        recovery_timeout: float = 60.0,
        expected_exception: Type[Exception] = Exception
    ):
        """
        Initialize circuit breaker
        
        Args:
            failure_threshold: Number of failures before opening circuit
            recovery_timeout: Time to wait before trying to close circuit
            expected_exception: Exception type that counts as failure
        """
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.expected_exception = expected_exception
        
        self.failure_count = 0
        self.last_failure_time = 0
        self.state = "closed"  # closed, open, half-open
        self.logger = logging.getLogger(__name__)
    
    async def call(self, func: Callable, *args, **kwargs) -> Any:
        """
        Execute function through circuit breaker
        
        Args:
            func: Function to execute
            *args: Positional arguments
            **kwargs: Keyword arguments
            
        Returns:
            Any: Function result
            
        Raises:
            Exception: If circuit is open or function fails
        """
        if self.state == "open":
            if asyncio.get_event_loop().time() - self.last_failure_time < self.recovery_timeout:
                raise Exception("Circuit breaker is open")
            else:
                self.state = "half-open"
                self.logger.info("Circuit breaker moved to half-open state")
        
        try:
            if asyncio.iscoroutinefunction(func):
                result = await func(*args, **kwargs)
            else:
                result = func(*args, **kwargs)
            
            # Success - reset failure count and close circuit
            if self.state == "half-open":
                self.state = "closed"
                self.logger.info("Circuit breaker closed after successful call")
            
            self.failure_count = 0
            return result
            
        except self.expected_exception as e:
            self.failure_count += 1
            self.last_failure_time = asyncio.get_event_loop().time()
            
            if self.failure_count >= self.failure_threshold:
                self.state = "open"
                self.logger.warning(f"Circuit breaker opened after {self.failure_count} failures")
            
            raise e
    
    def get_state(self) -> dict:
        """Get current circuit breaker state"""
        return {
            "state": self.state,
            "failure_count": self.failure_count,
            "failure_threshold": self.failure_threshold,
            "last_failure_time": self.last_failure_time,
            "recovery_timeout": self.recovery_timeout
        }