---
name: coding-standards
description: Coding standards, best practices, and patterns for Python development.
---

# Python Coding Standards & Best Practices

Coding standards for Python projects following modern best practices.

## Code Quality Principles

### 1. Readability First
- Code is read more than written
- Clear variable and function names
- Self-documenting code preferred over comments
- Consistent formatting

### 2. KISS (Keep It Simple, Stupid)
- Simplest solution that works
- Avoid over-engineering
- No premature optimization
- Easy to understand > clever code

### 3. DRY (Don't Repeat Yourself)
- Extract common logic into functions
- Create reusable components
- Share utilities across modules
- Avoid copy-paste programming

### 4. YAGNI (You Aren't Gonna Need It)
- Don't build features before they're needed
- Avoid speculative generality
- Add complexity only when required
- Start simple, refactor when needed

## Python Standards

### Variable Naming

```python
# ✅ GOOD: Descriptive snake_case names
search_query = "machine learning"
is_authenticated = True
total_count = 100
max_retries = 3

# ❌ BAD: Unclear names
q = "machine learning"
flag = True
x = 100
n = 3
```

### Function Naming

```python
# ✅ GOOD: Verb-noun pattern with snake_case
def fetch_user_data(user_id: str) -> dict:
    pass

def calculate_similarity(a: list[float], b: list[float]) -> float:
    pass

def is_valid_email(email: str) -> bool:
    pass

async def load_dataset(path: str) -> list[dict]:
    pass

# ❌ BAD: Unclear or noun-only
def data(id: str):
    pass

def similarity(a, b):
    pass

def email(e):
    pass
```

### Avoiding Unintended Mutations

```python
# ✅ GOOD: Don't mutate function arguments
def add_item_to_list(item: str, items: list[str]) -> list[str]:
    return items + [item]  # Returns new list

# ✅ ALSO GOOD: Be explicit about mutation
def add_item_in_place(item: str, items: list[str]) -> None:
    """Mutates the input list."""
    items.append(item)

# ❌ BAD: Unexpected mutation of arguments
def add_item(item: str, items: list[str]) -> list[str]:
    items.append(item)  # Mutates caller's list unexpectedly
    return items

# ✅ GOOD: For dataclasses, use replace when you need immutability
from dataclasses import dataclass, replace

@dataclass
class Config:
    timeout: int
    retries: int

updated_config = replace(config, timeout=60)

# Note: Mutation is fine in Python for performance when appropriate.
# Just avoid mutating function arguments unless it's the explicit purpose.
```

### Error Handling

```python
import json
import logging

logger = logging.getLogger(__name__)

# ✅ GOOD: Specific exception handling with proper error chaining
def load_config(file_path: str) -> dict:
    try:
        with open(file_path) as f:
            return json.load(f)
    except FileNotFoundError:
        logger.error(f"Config file not found: {file_path}")
        raise
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in config: {e}")
        raise ValueError(f"Failed to parse config file: {file_path}") from e

# ❌ BAD: No error handling
def load_config(file_path: str):
    with open(file_path) as f:
        return json.load(f)
```

### Exception Handling Best Practices

```python
# ✅ GOOD: Catch specific exceptions
try:
    result = process_data(data)
except ValueError as e:
    logger.error(f"Invalid data: {e}")
    raise
except KeyError as e:
    logger.error(f"Missing key: {e}")
    return default_value

# ❌ BAD: Catching generic Exception without good reason
try:
    result = process_data(data)
except Exception:
    pass  # Silently swallowing errors - NEVER do this

# ❌ BAD: Overly broad exception handling
try:
    result = process_data(data)
except Exception as e:
    # Only use this pattern when you truly need to catch ALL errors
    # (e.g., top-level error handlers, cleanup in finally blocks)
    logger.error(f"Unexpected error: {e}")
    raise

# ✅ ACCEPTABLE: Exception is fine for top-level handlers only
def main():
    try:
        run_application()
    except Exception as e:
        logger.critical(f"Application crashed: {e}")
        sys.exit(1)
```

**CRITICAL**: Never use bare `except Exception` to handle errors you should be handling specifically. If you can't name the specific exceptions that might occur, you don't understand your code well enough. Use `except Exception` only for:
- Top-level application error handlers
- Logging/monitoring wrappers
- Cleanup operations in finally blocks

### Async/Await Best Practices

```python
import asyncio

# ✅ GOOD: Parallel execution when possible
results, metadata, stats = await asyncio.gather(
    load_results(),
    load_metadata(),
    compute_stats()
)

# ❌ BAD: Sequential when unnecessary
results = await load_results()
metadata = await load_metadata()
stats = await compute_stats()

# ✅ GOOD: Handle errors in parallel tasks
results = await asyncio.gather(
    process_file("file1.txt"),
    process_file("file2.txt"),
    return_exceptions=True  # Returns exceptions instead of raising
)

for i, result in enumerate(results):
    if isinstance(result, Exception):
        logger.error(f"File {i+1} failed: {result}")
```

### Type Hints

```python
from typing import Literal
from datetime import datetime

# ✅ GOOD: Use native types (Python 3.9+)
def process_items(items: list[str]) -> dict[str, int]:
    return {item: len(item) for item in items}

def get_user_data(user_id: str) -> tuple[str, int] | None:
    pass

# ❌ BAD: Importing from typing when native types work
from typing import List, Dict, Tuple, Optional

def process_items(items: List[str]) -> Dict[str, int]:
    return {item: len(item) for item in items}

def get_user_data(user_id: str) -> Optional[Tuple[str, int]]:
    pass

# ✅ GOOD: Proper type hints with native types
from dataclasses import dataclass

@dataclass
class Document:
    id: str
    title: str
    content: str
    tags: list[str]

def get_document(doc_id: str) -> Document | None:
    # Implementation
    pass

# ❌ BAD: No type hints or using Any
from typing import Any

def get_document(doc_id: Any) -> Any:
    # Implementation
    pass
```

**Use native types when possible (Python 3.9+):**
- `list[T]` instead of `List[T]`
- `dict[K, V]` instead of `Dict[K, V]`
- `tuple[T, ...]` instead of `Tuple[T, ...]`
- `set[T]` instead of `Set[T]`
- `X | None` instead of `Optional[X]`
- `X | Y` instead of `Union[X, Y]`

## Pythonic Patterns

### String Formatting

```python
# ✅ GOOD: Use f-strings (Python 3.6+)
name = "Alice"
age = 30
message = f"User {name} is {age} years old"

# Also good for expressions
result = f"Sum: {a + b}, Product: {a * b}"

# ❌ BAD: Old-style formatting
message = "User %s is %d years old" % (name, age)
message = "User {} is {} years old".format(name, age)
```

### Context Managers

```python
# ✅ GOOD: Always use context managers for resources
with open("data.txt", "r") as f:
    data = f.read()

# Multiple context managers
with open("input.txt") as infile, open("output.txt", "w") as outfile:
    outfile.write(infile.read())

# ❌ BAD: Manual resource management
f = open("data.txt", "r")
data = f.read()
f.close()  # Easy to forget, won't run if exception occurs
```

### Path Handling

```python
from pathlib import Path

# ✅ GOOD: Use pathlib for modern path handling
data_dir = Path("data")
file_path = data_dir / "results.json"

if file_path.exists():
    content = file_path.read_text()

# Cross-platform paths
config_path = Path.home() / ".config" / "app" / "settings.json"

# ❌ BAD: Using os.path
import os

data_dir = "data"
file_path = os.path.join(data_dir, "results.json")

if os.path.exists(file_path):
    with open(file_path, "r") as f:
        content = f.read()
```

### Truthiness and None Checks

```python
# ✅ GOOD: Use truthiness for collections
if items:  # Check if list is non-empty
    process(items)

if not errors:  # Check if list is empty
    print("No errors")

# ✅ GOOD: Explicit None checks when None is different from empty
if value is not None:  # Distinguishes None from 0, "", [], etc.
    process(value)

if result is None:
    return default_value

# ❌ BAD: Explicit length checks for emptiness
if len(items) > 0:
    process(items)

# ❌ BAD: Using == for None
if value != None:  # Should use 'is not'
    process(value)
```

### Dictionary Operations

```python
# ✅ GOOD: Use get() with defaults
config = {"timeout": 30}
retry_count = config.get("retries", 3)  # Returns 3 if key missing

# ✅ GOOD: Use setdefault() for initialization
cache = {}
cache.setdefault("results", []).append(new_result)

# ✅ GOOD: Use defaultdict for repeated defaults
from collections import defaultdict

scores = defaultdict(int)  # Default value is 0
scores["alice"] += 10  # No KeyError

# ✅ GOOD: Dictionary comprehensions
squared = {x: x**2 for x in range(5)}
filtered = {k: v for k, v in data.items() if v > 0}

# ❌ BAD: Manual key checking
if "retries" in config:
    retry_count = config["retries"]
else:
    retry_count = 3
```

### Iteration Patterns

```python
# ✅ GOOD: Use enumerate() for index + value
for i, item in enumerate(items):
    print(f"{i}: {item}")

# Start from different index
for i, item in enumerate(items, start=1):
    print(f"Item {i}: {item}")

# ✅ GOOD: Use zip() to iterate multiple sequences
names = ["Alice", "Bob"]
ages = [30, 25]
for name, age in zip(names, ages):
    print(f"{name} is {age}")

# ✅ GOOD: Unpack sequences
x, y, z = coordinates
first, *rest = items  # first = items[0], rest = items[1:]
*others, last = items  # last = items[-1], others = items[:-1]

# ❌ BAD: Manual indexing
for i in range(len(items)):
    item = items[i]
    print(f"{i}: {item}")
```

### Logging Best Practices

```python
import logging

# ✅ GOOD: Use logger, not print
logger = logging.getLogger(__name__)

def process_data(data):
    logger.info(f"Processing {len(data)} items")
    try:
        result = expensive_operation(data)
        logger.debug(f"Result: {result}")
        return result
    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        raise

# ✅ GOOD: Use appropriate log levels
logger.debug("Detailed information for debugging")
logger.info("General informational messages")
logger.warning("Warning messages for potentially harmful situations")
logger.error("Error messages for serious problems")
logger.critical("Critical messages for very serious errors")

# ❌ BAD: Using print for logging
def process_data(data):
    print(f"Processing {len(data)} items")  # Lost in production
    result = expensive_operation(data)
    print(f"Result: {result}")
    return result
```

## File Organization

### Project Structure

```
project/
├── src/                   # Source code
│   ├── core/             # Core business logic
│   ├── utils/            # Utility functions
│   ├── models/           # Data models/classes
│   └── config.py         # Configuration
├── tests/
│   ├── unit/             # Unit tests
│   ├── integration/      # Integration tests
│   └── conftest.py       # Pytest configuration
├── scripts/              # Standalone scripts
├── data/                 # Data files (if applicable)
├── pyproject.toml        # Project dependencies (uv/poetry)
├── README.md
└── .gitignore
```

### File Naming

```
models/user.py               # snake_case for modules
utils/text_processor.py      # snake_case for modules
core/analyzer.py             # snake_case for modules
tests/test_analyzer.py       # test_ prefix for test files
```

### Module Organization

```python
# ✅ GOOD: Clear module organization with proper imports
# models/user.py
from dataclasses import dataclass
from datetime import datetime

@dataclass
class User:
    id: str
    name: str
    created_at: datetime

# utils/text_processor.py
def normalize_text(text: str) -> str:
    """Normalize text by lowercasing and removing extra spaces."""
    return " ".join(text.lower().split())

def tokenize(text: str) -> list[str]:
    """Split text into tokens."""
    return text.split()

# core/analyzer.py
from models.user import User
from utils.text_processor import normalize_text

def analyze_user_data(user: User, text: str) -> dict[str, int]:
    """Analyze text data for a user."""
    normalized = normalize_text(text)
    return {
        "user_id": user.id,
        "word_count": len(normalized.split())
    }
```

## Comments & Documentation

### When to Comment

```python
# ✅ GOOD: Explain WHY, not WHAT
# Use exponential backoff to avoid overwhelming the API during outages
delay = min(1000 * (2 ** retry_count), 30000)

# Deliberately using mutation here for performance with large arrays
items.append(new_item)

# ❌ BAD: Stating the obvious
# Increment counter by 1
count += 1

# Set name to user's name
name = user.name
```

### Docstrings for Public APIs

```python
def search_documents(
    query: str,
    limit: int = 10,
    min_score: float = 0.5
) -> list[dict]:
    """Search documents using text similarity.

    Args:
        query: Search query string
        limit: Maximum number of results. Defaults to 10.
        min_score: Minimum similarity score (0-1). Defaults to 0.5.

    Returns:
        List of matching documents sorted by similarity score

    Raises:
        ValueError: If query is empty or min_score is out of range

    Example:
        >>> results = search_documents("python testing", limit=5)
        >>> print(results[0]["title"])
        "Introduction to pytest"
    """
    if not query:
        raise ValueError("Query cannot be empty")
    if not 0 <= min_score <= 1:
        raise ValueError("min_score must be between 0 and 1")
    # Implementation
    pass
```

### Docstring Styles

```python
# ✅ GOOD: Google-style docstrings (recommended)
def calculate_similarity(vector_a: list[float], vector_b: list[float]) -> float:
    """Calculate cosine similarity between two vectors.

    Args:
        vector_a: First vector
        vector_b: Second vector

    Returns:
        Similarity score between -1 and 1
    """
    pass

# ✅ ALSO GOOD: NumPy-style docstrings (for scientific code)
def calculate_similarity(vector_a: list[float], vector_b: list[float]) -> float:
    """Calculate cosine similarity between two vectors.

    Parameters
    ----------
    vector_a : list[float]
        First vector
    vector_b : list[float]
        Second vector

    Returns
    -------
    float
        Similarity score between -1 and 1
    """
    pass
```

## Performance Best Practices

### Caching with functools

```python
from functools import lru_cache, cache

# ✅ GOOD: Cache expensive computations
@lru_cache(maxsize=128)
def calculate_similarity(vector_a: tuple[float, ...], vector_b: tuple[float, ...]) -> float:
    # Expensive computation
    pass

# For Python 3.9+: Simple unbounded cache
@cache
def get_config(key: str) -> str:
    # Expensive I/O operation
    pass

# ❌ BAD: Recomputing the same thing repeatedly
def calculate_similarity(vector_a: list[float], vector_b: list[float]) -> float:
    # This will recompute every time, even for same inputs
    pass
```

### List Comprehensions vs Loops

```python
# ✅ GOOD: List comprehensions are faster
squared = [x**2 for x in numbers if x > 0]

# ❌ SLOWER: Traditional loop
squared = []
for x in numbers:
    if x > 0:
        squared.append(x**2)

# ✅ GOOD: Generator for large datasets
squared_gen = (x**2 for x in numbers if x > 0)

# ✅ GOOD: Use sum() with generator for memory efficiency
total = sum(x**2 for x in large_dataset if x > 0)
```

### Built-in Functions

```python
# ✅ GOOD: Use built-ins (they're optimized in C)
numbers = [1, 2, 3, 4, 5]
total = sum(numbers)
maximum = max(numbers)
minimum = min(numbers)

# ✅ GOOD: Use any() and all()
has_errors = any(item.is_error for item in results)
all_valid = all(item.is_valid for item in results)

# ❌ SLOWER: Manual loops
total = 0
for num in numbers:
    total += num
```

## Testing Standards (pytest)

### Test Structure (AAA Pattern)

```python
def test_calculates_similarity_correctly():
    # Arrange
    vector1 = [1.0, 0.0, 0.0]
    vector2 = [0.0, 1.0, 0.0]

    # Act
    similarity = calculate_cosine_similarity(vector1, vector2)

    # Assert
    assert similarity == 0.0


# ✅ GOOD: Use fixtures for setup
import pytest
from models.document import Document

@pytest.fixture
def sample_document():
    return Document(
        id="123",
        title="Test Document",
        content="Sample content",
        tags=["test"]
    )

def test_document_creation(sample_document):
    assert sample_document.id == "123"
    assert sample_document.title == "Test Document"
    assert len(sample_document.tags) == 1
```

### Test Naming

```python
# ✅ GOOD: Descriptive test names with test_ prefix
def test_returns_empty_list_when_no_documents_match_query():
    pass

def test_raises_error_when_file_not_found():
    pass

def test_normalizes_text_by_lowercasing_and_removing_extra_spaces():
    pass

# ❌ BAD: Vague test names
def test_works():
    pass

def test_search():
    pass
```

### Async Tests

```python
import pytest

# ✅ GOOD: Testing async functions
@pytest.mark.asyncio
async def test_load_dataset():
    # Arrange
    path = "data/test_dataset.json"

    # Act
    result = await load_dataset(path)

    # Assert
    assert result is not None
    assert len(result) > 0
    assert "id" in result[0]
```

### Mocking

```python
from unittest.mock import Mock, patch, mock_open

# ✅ GOOD: Mock file operations
@patch("builtins.open", mock_open(read_data='{"key": "value"}'))
def test_load_config_reads_file():
    config = load_config("config.json")
    assert config["key"] == "value"

# ✅ GOOD: Mock external dependencies
@patch("utils.processor.external_api_call")
def test_process_data_calls_external_api(mock_api):
    # Arrange
    mock_api.return_value = {"status": "success"}

    # Act
    result = process_data({"input": "test"})

    # Assert
    mock_api.assert_called_once()
    assert result["status"] == "success"

# ✅ GOOD: Mock with side effects for error testing
def test_handles_file_not_found():
    with patch("builtins.open", side_effect=FileNotFoundError):
        with pytest.raises(FileNotFoundError):
            load_config("missing.json")
```

## Code Smell Detection

Watch for these anti-patterns:

### 1. Long Functions
```python
# ❌ BAD: Function > 50 lines
def process_dataset(data):
    # 100 lines of code
    pass

# ✅ GOOD: Split into smaller functions
def process_dataset(data):
    validated = validate_dataset(data)
    transformed = transform_dataset(validated)
    return save_dataset(transformed)
```

### 2. Deep Nesting
```python
# ❌ BAD: 5+ levels of nesting
if user:
    if user.is_authenticated:
        if document:
            if document.is_accessible:
                if has_permission:
                    # Do something
                    pass

# ✅ GOOD: Early returns
def process_document(user, document, has_permission):
    if not user:
        return
    if not user.is_authenticated:
        return
    if not document:
        return
    if not document.is_accessible:
        return
    if not has_permission:
        return

    # Do something
    pass
```

### 3. Magic Numbers
```python
# ❌ BAD: Unexplained numbers
if retry_count > 3:
    pass
time.sleep(0.5)

# ✅ GOOD: Named constants
MAX_RETRIES = 3
DEBOUNCE_DELAY_SECONDS = 0.5

if retry_count > MAX_RETRIES:
    pass
time.sleep(DEBOUNCE_DELAY_SECONDS)
```

### 4. Mutable Default Arguments
```python
# ❌ BAD: Mutable default argument
def add_item(item, items=[]):
    items.append(item)
    return items

# ✅ GOOD: Use None as default
def add_item(item, items=None):
    if items is None:
        items = []
    items.append(item)
    return items
```

### 5. String Concatenation in Loops
```python
# ❌ BAD: String concatenation in loop
result = ""
for item in items:
    result += str(item) + ","

# ✅ GOOD: Use join
result = ",".join(str(item) for item in items)
```

### 6. Complex Comprehensions
```python
# ❌ BAD: Unreadable nested comprehension
result = [
    item.upper()
    for sublist in data
    for item in sublist
    if item
    if len(item) > 3
    if not item.startswith("_")
]

# ✅ GOOD: Break into multiple steps for readability
flattened = (item for sublist in data for item in sublist)
filtered = (item for item in flattened if item and len(item) > 3)
result = [item.upper() for item in filtered if not item.startswith("_")]

# ✅ ALSO GOOD: Use a regular loop when it's clearer
result = []
for sublist in data:
    for item in sublist:
        if item and len(item) > 3 and not item.startswith("_"):
            result.append(item.upper())
```

**Remember**: Code quality is not negotiable. Clear, maintainable code enables rapid development and confident refactoring.