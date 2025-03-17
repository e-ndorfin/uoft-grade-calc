from dataclasses import dataclass
from typing import List

@dataclass
class CategoryConfig:
    name: str
    weight: float
    total_items: int
    best_of: int

@dataclass
class ClassConfig:
    class_name: str
    categories: List[CategoryConfig]
