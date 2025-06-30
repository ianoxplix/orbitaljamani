# models.py

from dataclasses import dataclass, field
from typing import List

@dataclass
class QuoteMeta:
    quotation_no: str
    date: str
    order_no: str
    service: str  # ✅ Add this

@dataclass
class Customer:
    name: str
    address: str
    town: str
    contact: str
    phone: str
    email: str

@dataclass
class QuoteItem:
    qty: float
    description: str
    units: str
    unit_price: float
    total: float

@dataclass
class Quote:
    customer: Customer
    quote: QuoteMeta
    items: List[QuoteItem]
    labour: QuoteItem
    grand_total: float

