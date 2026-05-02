# Kiyabo Duka — Migration to Python Web Stack
## Full Architecture Analysis, Django Monolith Design & FastAPI+Next.js Comparison

**Based on:** Kiyabo Duka v0.031 — 33 tables, 4,020+ sales records, retail shop system  
**Your profile:** Pythonic developer, considering long-term platform independence  
**Prepared:** May 2026  

---

## Table of Contents

1. [The Honest Migration Assessment](#1-the-honest-migration-assessment)
2. [Why Django Is the Right Choice for You](#2-why-django-is-the-right-choice-for-you)
3. [FastAPI + Next.js — When It Makes Sense](#3-fastapi--nextjs--when-it-makes-sense)
4. [Head-to-Head Comparison](#4-head-to-head-comparison)
5. [Django Project Architecture — Full Design](#5-django-project-architecture--full-design)
6. [Data Models — Access to Django ORM](#6-data-models--access-to-django-orm)
7. [Django Admin — Your Free frm_master](#7-django-admin--your-free-frm_master)
8. [Views, URLs & Business Logic](#8-views-urls--business-logic)
9. [Reports in Django — WeasyPrint + Charts](#9-reports-in-django--weasyprint--charts)
10. [Migration Strategy — Access to PostgreSQL](#10-migration-strategy--access-to-postgresql)
11. [Deployment — Production-Ready Setup](#11-deployment--production-ready-setup)
12. [FastAPI + Next.js Deep Dive](#12-fastapi--nextjs-deep-dive)
13. [The Decision Framework](#13-the-decision-framework)
14. [Complete Django Boilerplate](#14-complete-django-boilerplate)

---

## 1. The Honest Migration Assessment

Before diving into architecture, let's be completely honest about where you are and what a migration actually means.

### What You Have in Access Today

Your Access database is not a simple flat-file system. It is a legitimately well-architected relational system:

- **33 tables** with proper normalization (categories → types → products → specs is textbook)
- **44 enforced relationships** with referential integrity
- **9 VBA modules** containing real business logic (stock CRUD, accounting calculations, expense generation, rate calculations)
- **4,020+ sales transactions** — live production data
- **191 debt return records**, **357 purchase details** — financial history that cannot be lost

The VBA modules are the most important thing to understand before migrating. They are not just UI code — `modAccountings`, `modExpenses`, `modPayables`, `modRATE`, `modStats` contain **business rules**. Every single one of those needs to be re-expressed in Python before you can consider migration complete. This is the real work.

### What Migration Actually Costs

| Work Item | Complexity | Time Estimate |
|---|---|---|
| Schema → Django models | Low (1:1 mapping) | 1–2 days |
| Data migration (Access → PostgreSQL) | Medium | 1 day |
| VBA business logic → Python services | **High** | 1–2 weeks |
| Forms → Django views/templates | Medium | 1–2 weeks |
| Reports → WeasyPrint/PDF | Medium | 3–5 days |
| Testing & QA | High | 1 week |
| Deployment & production setup | Medium | 2–3 days |
| **TOTAL** | | **4–6 weeks minimum** |

### The Key Question

**Should you migrate now, or finish the Access system first?**

Honest answer: **Finish the Access system first.** Here is why:

1. Your Access data model is sound — it will map cleanly to Django models
2. The 12-day plan to complete Access forms/reports is real — it takes less time than migrating
3. A completed Access system gives you a **working reference** when building Django views
4. You will make better architectural decisions in Django after you fully understand the business logic by building it in VBA first
5. Access gives you a working system TODAY. Django gives you one in 6 weeks.

**The right sequence:**
```
Phase 1 (Now):      Complete Access system (12 days)
Phase 2 (Month 2):  Build Django in parallel (Access still running)
Phase 3 (Month 3):  Migrate data, switch over, retire Access
```

This document is your Phase 2 blueprint. You can start designing the Django structure now while finishing Access — they are not mutually exclusive.

---

## 2. Why Django Is the Right Choice for You

### You Are Pythonic — Django Is the Python-Native Full Stack

Django is not just a web framework. It is the closest thing Python has to a complete, batteries-included application platform — which is exactly what a retail shop management system needs.

**What Django gives you out of the box that directly replaces Access features:**

| Access Feature | Django Equivalent | Effort |
|---|---|---|
| frm_master (dashboard) | Django Admin + custom views | Low |
| Form wizards & data entry | Django ModelForms + Class-Based Views | Low |
| Combo boxes (FK lookups) | `ForeignKey` + `ModelChoiceField` | Zero — automatic |
| Subforms (e.g. purchase + details) | Django inline formsets | Medium |
| Reports (Access Reports) | WeasyPrint / ReportLab / xhtml2pdf | Medium |
| Built-in user auth | `django.contrib.auth` | Zero — built-in |
| Data validation | Django model validators + form validation | Low |
| Data migrations | `manage.py makemigrations && migrate` | Zero — automatic |
| Multi-user access | Built-in (HTTP is naturally multi-user) | Zero |
| Run on any OS | Yes — Linux, Mac, Windows, Raspberry Pi | Zero |
| Web accessible (mobile, tablet) | Yes — browser-based | Zero |
| Backup | `pg_dump` or `manage.py dumpdata` | Trivial |

### The Deepest Django Advantage: The ORM

Your Access queries translated to Django ORM are more readable, safer, and refactorable. Compare:

**Access SQL (ageing query — from your schema):**
```sql
SELECT d.debtor_name, d.phone_number_1,
    SUM(IIf(DateDiff("d", dt.expected_payment_date, Date()) <= 0, 
        dt.amount - Nz(paid.total_paid,0), 0)) AS current_balance
FROM tbl_debtors d INNER JOIN tbl_debts dt ON d.debtor_id = dt.debtor_id
LEFT JOIN (SELECT debt_id, SUM(amount) AS total_paid 
           FROM tbl_debt_returns GROUP BY debt_id) AS paid 
ON dt.sale_id = paid.debt_id
GROUP BY d.debtor_id, d.debtor_name, d.phone_number_1
```

**Django ORM equivalent:**
```python
from django.db.models import Sum, F, ExpressionWrapper, DecimalField
from django.db.models.functions import Coalesce
from django.utils import timezone

today = timezone.now().date()

debtors_with_balances = (
    Debtor.objects
    .annotate(
        total_debt=Coalesce(Sum('debts__amount'), 0),
        total_paid=Coalesce(Sum('debts__debt_returns__amount'), 0),
    )
    .annotate(
        outstanding=ExpressionWrapper(
            F('total_debt') - F('total_paid'),
            output_field=DecimalField()
        )
    )
    .filter(outstanding__gt=0)
    .order_by('-outstanding')
)
```

This is Python. You can `import` it, test it, put it in a service class, call it from a view or a management command or a scheduled task — no VBA, no Access, no Windows dependency.

### Django Admin: The Fastest Win

Here is the thing most people miss: **Django Admin is a fully functional CRUD interface that generates itself from your models.** For your Access system:

- Register all 33 models in `admin.py`
- Customize `list_display`, `search_fields`, `list_filter`
- Add `InlineAdmin` for parent-child relationships (purchases + purchase details, etc.)
- **You get frm_master, frm_products_list, frm_purchases, frm_suppliers for free in about 2 hours**

For a single-location shop with a handful of staff, Django Admin IS a viable production interface. You then build custom views only for the things that need real UX polish (the POS sales entry form, the debtor payment screen).

---

## 3. FastAPI + Next.js — When It Makes Sense

FastAPI + Next.js (or any React frontend) is not a wrong choice — it is a different trade-off. You need to understand exactly what you are signing up for.

### What This Stack Is

- **FastAPI** — async Python REST API framework. Extremely fast. Auto-generates OpenAPI docs. Built on Pydantic + Starlette.
- **Next.js** — React framework with server-side rendering, file-based routing, and API routes.
- Together they form a **decoupled architecture**: the backend is a pure JSON API, the frontend is a separate JavaScript application.

### What This Stack Is NOT

It is not a form builder. It is not a report generator. It does not come with an admin panel. It does not have database migrations built in (you add Alembic). It does not have user authentication built in (you add it). Every single thing that Django gives you for free is something you have to assemble yourself.

For a shop management system with forms, reports, role-based access, financial data, and admin operations — that is a lot of assembly.

### Where FastAPI + Next.js Would Make Real Sense for Kiyabo Duka

Specifically and only in these cases:

1. **You want a mobile app** (React Native can share frontend code with Next.js)
2. **Multiple branches** — and you need a single API serving a web app, a mobile app, and possibly third-party integrations (accounting software, M-Pesa API, etc.)
3. **Real-time features** — live sales dashboard updating without page refresh (WebSocket-based, using FastAPI's async support)
4. **You plan to hire a frontend developer** who knows React but not Django templates
5. **The system grows to enterprise scale** and you need to independently scale the API tier from the frontend tier

For a single-shop system managed by one developer (you), this is over-engineering. You would spend 40% of your time writing JavaScript that Django would have handled automatically.

---

## 4. Head-to-Head Comparison

### The Definitive Comparison for Your Specific System

| Dimension | Django Monolith | FastAPI + Next.js |
|---|---|---|
| **Setup time to first working form** | 30 minutes | 3–4 hours |
| **Database migrations** | Built-in (`makemigrations`) | Manual (Alembic) |
| **Admin/backoffice interface** | Free (Django Admin) | Build from scratch |
| **Form handling & validation** | Built-in (ModelForms) | Manual (Pydantic + React Hook Form) |
| **Authentication** | Built-in (`django.contrib.auth`) | Assemble (JWT + python-jose) |
| **PDF reports** | WeasyPrint + templates | Same (WeasyPrint) |
| **ORM quality** | Django ORM (excellent) | SQLAlchemy (excellent, more verbose) |
| **Your comfort level (Pythonic)** | Very high | High backend, low frontend |
| **Multi-user / concurrent** | Good (async via ASGI) | Excellent (native async) |
| **Mobile app in future** | Possible (add DRF later) | Native fit |
| **Codebase size for same features** | Small (Django handles magic) | 2–3× larger |
| **Learning curve** | Low — Django docs are world-class | Medium — two separate stacks |
| **Deployment complexity** | Simple (1 process + gunicorn) | Complex (2 processes + reverse proxy) |
| **Real-time dashboards** | Django Channels (adds complexity) | FastAPI WebSockets (native) |
| **Community size** | Massive | Growing fast |
| **Job market relevance** | High | High + growing |
| **Total dev time for your system** | ~4 weeks | ~7–9 weeks |

### The Verdict for Kiyabo Duka v1.0

**Choose Django.** Here is the reasoning, specific to your system:

- You have 33 tables of CRUD. Django ModelForms + Class-Based Views handle CRUD magnificently with minimal code.
- You have 9 VBA modules of business logic. These become Django service layer classes — pure Python, perfectly at home in Django.
- You have financial reports (P&L, cash flow, balance sheet). These are SQL + templates — Django's template engine renders them to HTML, WeasyPrint converts to PDF. This works identically in both stacks, but Django has less wiring.
- You are one developer. Maintaining one Python codebase is half the maintenance of a Python API + JavaScript frontend.
- Django Admin covers your entire backoffice in 2 hours. The sales POS is the only screen that needs real UX work.

**Add FastAPI later as a companion service if and when you need:**
- An M-Pesa payment webhook endpoint (FastAPI is ideal for webhook receivers)
- A mobile app API (add Django REST Framework to your Django app, or a small FastAPI service)
- Real-time sales dashboard (Django Channels or a FastAPI WebSocket endpoint)

The hybrid approach — Django core + FastAPI for specific async/real-time endpoints — is actually common in production and works well.

---

## 5. Django Project Architecture — Full Design

### Project Structure

```
kiyabo_duka/                        ← git root
├── manage.py
├── requirements.txt
├── requirements-dev.txt
├── .env                            ← never commit this
├── .env.example
├── pytest.ini
│
├── config/                         ← Django project settings package
│   ├── __init__.py
│   ├── settings/
│   │   ├── base.py                 ← shared settings
│   │   ├── development.py          ← local dev overrides
│   │   └── production.py           ← production overrides
│   ├── urls.py                     ← root URL conf
│   ├── wsgi.py
│   └── asgi.py
│
├── apps/                           ← all Django apps live here
│   ├── core/                       ← shared utilities, base models, mixins
│   │   ├── models.py               ← TimeStampedModel, SoftDeleteModel
│   │   ├── mixins.py
│   │   └── utils.py
│   │
│   ├── catalog/                    ← products, categories, specs, brands, units
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   ├── services.py             ← business logic (no DB in views)
│   │   └── tests/
│   │
│   ├── inventory/                  ← purchases, stock levels, returns outward
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   ├── services.py
│   │   └── tests/
│   │
│   ├── sales/                      ← sales, office use, return inwards
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   ├── services.py             ← stock decrement, receipt generation
│   │   └── tests/
│   │
│   ├── credit/                     ← debtors, debts, debt returns
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   ├── services.py             ← balance calculation, ageing
│   │   └── tests/
│   │
│   ├── finance/                    ← expenses, payments, obligations, prepayments
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   ├── services.py             ← modExpenses, modPayables logic
│   │   └── tests/
│   │
│   ├── assets/                     ← fixed assets, liabilities, drawings
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── forms.py
│   │   ├── urls.py
│   │   ├── admin.py
│   │   └── tests/
│   │
│   ├── reports/                    ← all report generation
│   │   ├── views.py                ← report hub + PDF endpoints
│   │   ├── urls.py
│   │   ├── generators/
│   │   │   ├── sales.py            ← daily, monthly, top products
│   │   │   ├── stock.py            ← stock levels, low stock
│   │   │   ├── credit.py           ← debtor ageing, collections
│   │   │   ├── finance.py          ← P&L, cash flow, expenses
│   │   │   └── balance_sheet.py
│   │   └── templates/
│   │       └── reports/            ← HTML templates for PDF rendering
│   │
│   └── dashboard/                  ← frm_master equivalent
│       ├── views.py
│       ├── urls.py
│       └── templates/
│           └── dashboard/
│
├── templates/                      ← global templates
│   ├── base.html                   ← main layout (navbar, sidebar, footer)
│   ├── base_print.html             ← print/PDF layout
│   └── components/                 ← reusable partials
│       ├── _table.html
│       ├── _form_errors.html
│       ├── _pagination.html
│       └── _kpi_card.html
│
├── static/                         ← CSS, JS, images
│   ├── css/
│   │   └── main.css                ← Tailwind CSS output (or custom)
│   ├── js/
│   │   └── main.js
│   └── img/
│       └── logo.png
│
└── media/                          ← user uploads (product images etc.)
```

### Why This Structure

- **One app per business domain** — not one app per database table. `catalog` owns everything about products; `credit` owns everything about debt.
- **`services.py` in each app** — this is where your VBA module logic goes. Views never contain business logic. Services never know about HTTP.
- **`core` app** — base model classes (`TimeStampedModel` with `created_at`/`updated_at`) shared across all apps.
- **`reports` app** — completely separate. Report generation is a read-only operation — it belongs apart from CRUD apps.
- **Split settings** — `base.py` + `development.py` + `production.py`. This is standard Django best practice and saves you from committing secrets.

---

## 6. Data Models — Access to Django ORM

### Base Model (apps/core/models.py)

```python
# apps/core/models.py
from django.db import models


class TimeStampedModel(models.Model):
    """
    Abstract base class giving every model created_at / updated_at.
    Access equivalent: adding created_by / modified_date columns to all tables.
    """
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        abstract = True


class SoftDeleteModel(TimeStampedModel):
    """
    Soft delete — sets is_deleted instead of actual DELETE.
    Useful for financial records that must never truly disappear.
    """
    is_deleted = models.BooleanField(default=False)
    deleted_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        abstract = True
```

### Catalog Models (apps/catalog/models.py)

```python
# apps/catalog/models.py
from django.db import models
from apps.core.models import TimeStampedModel


class Category(TimeStampedModel):
    """tbl_categories"""
    name = models.CharField(max_length=255, unique=True)

    class Meta:
        db_table = 'catalog_category'
        verbose_name_plural = 'categories'
        ordering = ['name']

    def __str__(self):
        return self.name


class ProductType(TimeStampedModel):
    """tbl_types"""
    category = models.ForeignKey(
        Category, on_delete=models.PROTECT, related_name='types'
    )
    name = models.CharField(max_length=255)

    class Meta:
        db_table = 'catalog_product_type'
        unique_together = [('category', 'name')]
        ordering = ['category__name', 'name']

    def __str__(self):
        return f"{self.category.name} — {self.name}"


class Brand(TimeStampedModel):
    """tbl_brands"""
    name = models.CharField(max_length=255, unique=True)

    class Meta:
        db_table = 'catalog_brand'
        ordering = ['name']

    def __str__(self):
        return self.name


class Unit(TimeStampedModel):
    """tbl_units"""
    name = models.CharField(max_length=255)
    abbreviation = models.CharField(max_length=20)

    class Meta:
        db_table = 'catalog_unit'

    def __str__(self):
        return f"{self.name} ({self.abbreviation})"


class Spec(TimeStampedModel):
    """tbl_specs — e.g. 'Size', 'Color', 'Weight'"""
    name = models.CharField(max_length=255)

    class Meta:
        db_table = 'catalog_spec'

    def __str__(self):
        return self.name


class SpecValue(TimeStampedModel):
    """tbl_spec_values — e.g. 'Large', 'Blue', '500g'"""
    spec = models.ForeignKey(Spec, on_delete=models.PROTECT, related_name='values')
    value = models.CharField(max_length=255)

    class Meta:
        db_table = 'catalog_spec_value'
        unique_together = [('spec', 'value')]

    def __str__(self):
        return f"{self.spec.name}: {self.value}"


class Product(TimeStampedModel):
    """tbl_products"""
    name = models.CharField(max_length=255)
    product_type = models.ForeignKey(
        ProductType, on_delete=models.PROTECT, related_name='products'
    )
    brand = models.ForeignKey(
        Brand, on_delete=models.SET_NULL, null=True, blank=True, related_name='products'
    )
    unit = models.ForeignKey(
        Unit, on_delete=models.SET_NULL, null=True, blank=True
    )
    image = models.ImageField(upload_to='products/', null=True, blank=True)
    barcode = models.CharField(max_length=100, blank=True, db_index=True)

    class Meta:
        db_table = 'catalog_product'
        unique_together = [('name', 'product_type', 'brand')]
        ordering = ['name']

    def __str__(self):
        return self.name


class ProductSpec(TimeStampedModel):
    """
    tbl_product_specs — a specific variant of a product.
    Example: 'Coca Cola' (product) + '500ml' (spec_value) = one ProductSpec.
    This is the unit that gets bought, sold, and stocked.
    """
    product = models.ForeignKey(
        Product, on_delete=models.PROTECT, related_name='specs'
    )
    spec_value = models.ForeignKey(
        SpecValue, on_delete=models.PROTECT, related_name='product_specs'
    )
    default_cost_price = models.DecimalField(
        max_digits=12, decimal_places=2, null=True, blank=True
    )
    default_selling_price = models.DecimalField(
        max_digits=12, decimal_places=2, null=True, blank=True
    )
    reorder_level = models.PositiveIntegerField(default=5)
    current_stock = models.IntegerField(default=0)

    class Meta:
        db_table = 'catalog_product_spec'
        unique_together = [('product', 'spec_value')]

    def __str__(self):
        return f"{self.product.name} — {self.spec_value.value}"

    @property
    def is_low_stock(self) -> bool:
        return self.current_stock <= self.reorder_level

    @property
    def stock_value(self):
        """Current inventory value at cost price."""
        if self.default_cost_price:
            return self.current_stock * self.default_cost_price
        return 0
```

### Sales Models (apps/sales/models.py)

```python
# apps/sales/models.py
from django.db import models
from django.utils import timezone
from apps.core.models import TimeStampedModel
from apps.catalog.models import ProductSpec


class PaymentMethod(TimeStampedModel):
    """tbl_payment_methods"""
    name = models.CharField(max_length=255, unique=True)

    class Meta:
        db_table = 'sales_payment_method'

    def __str__(self):
        return self.name


class SaleReceipt(TimeStampedModel):
    """
    NEW — groups multiple Sale items into one transaction.
    This solves the missing receipt_id issue from the Access schema.
    """
    receipt_number = models.CharField(max_length=50, unique=True, db_index=True)
    sale_date = models.DateTimeField(default=timezone.now)
    payment_method = models.ForeignKey(
        PaymentMethod, on_delete=models.PROTECT, related_name='receipts'
    )
    notes = models.TextField(blank=True)

    class Meta:
        db_table = 'sales_receipt'
        ordering = ['-sale_date']

    def __str__(self):
        return self.receipt_number

    @property
    def total_amount(self):
        return sum(item.amount for item in self.items.all())

    @classmethod
    def generate_receipt_number(cls) -> str:
        from django.utils import timezone
        today = timezone.now().strftime('%Y%m%d')
        count = cls.objects.filter(
            receipt_number__startswith=f'RCP-{today}'
        ).count() + 1
        return f'RCP-{today}-{count:03d}'


class Sale(TimeStampedModel):
    """
    tbl_sales — one line item in a sale.
    Linked to SaleReceipt for grouping.
    The original tbl_sales had no receipt grouping — this fixes it.
    """
    receipt = models.ForeignKey(
        SaleReceipt, on_delete=models.PROTECT,
        related_name='items', null=True, blank=True
    )
    product_spec = models.ForeignKey(
        ProductSpec, on_delete=models.PROTECT, related_name='sales'
    )
    quantity = models.PositiveIntegerField(default=1)
    unit_price = models.DecimalField(max_digits=12, decimal_places=2)
    discount = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    sale_date = models.DateTimeField(default=timezone.now, db_index=True)
    payment_method = models.ForeignKey(
        PaymentMethod, on_delete=models.PROTECT,
        related_name='sales', null=True, blank=True
    )
    notes = models.TextField(blank=True)

    class Meta:
        db_table = 'sales_sale'
        ordering = ['-sale_date']

    def __str__(self):
        return f"Sale #{self.pk} — {self.product_spec}"

    @property
    def amount(self):
        return (self.quantity * self.unit_price) - self.discount

    def save(self, *args, **kwargs):
        is_new = self.pk is None
        super().save(*args, **kwargs)
        if is_new:
            # Decrement stock on save — equivalent to modStockCRUD
            from apps.inventory.services import StockService
            StockService.decrease_stock(self.product_spec, self.quantity)


class OfficeUse(TimeStampedModel):
    """tbl_office_use + tbl_sales_office_use merged"""
    OFFICE_USE_CHOICES = [
        ('personal', 'Personal Use'),
        ('business', 'Business Expense'),
        ('sample', 'Sample/Promotion'),
        ('damaged', 'Damaged Goods'),
        ('other', 'Other'),
    ]
    product_spec = models.ForeignKey(
        ProductSpec, on_delete=models.PROTECT, related_name='office_uses'
    )
    use_type = models.CharField(max_length=50, choices=OFFICE_USE_CHOICES)
    quantity = models.PositiveIntegerField(default=1)
    unit_price = models.DecimalField(max_digits=12, decimal_places=2)
    sale_date = models.DateTimeField(default=timezone.now)
    reason = models.TextField(blank=True)

    class Meta:
        db_table = 'sales_office_use'

    @property
    def amount(self):
        return self.quantity * self.unit_price
```

### Credit Models (apps/credit/models.py)

```python
# apps/credit/models.py
from django.db import models
from django.utils import timezone
from django.db.models import Sum
from django.db.models.functions import Coalesce
from apps.core.models import TimeStampedModel
from apps.catalog.models import ProductSpec
from apps.sales.models import PaymentMethod


class Debtor(TimeStampedModel):
    """tbl_debtors"""
    name = models.CharField(max_length=255, db_index=True)
    address = models.TextField(blank=True)
    phone_1 = models.CharField(max_length=20, blank=True)
    phone_2 = models.CharField(max_length=20, blank=True)
    nida_id = models.CharField(max_length=255, blank=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        db_table = 'credit_debtor'
        ordering = ['name']

    def __str__(self):
        return self.name

    @property
    def outstanding_balance(self):
        """
        Total owed minus total paid.
        Equivalent to the balance calculation in VBA modAccountings.
        """
        total_debt = self.debts.aggregate(
            total=Coalesce(Sum('amount'), 0)
        )['total']
        total_paid = DebtReturn.objects.filter(
            debt__debtor=self
        ).aggregate(total=Coalesce(Sum('amount'), 0))['total']
        return total_debt - total_paid

    @property
    def has_overdue(self):
        return self.debts.filter(
            expected_payment_date__lt=timezone.now().date()
        ).exists()


class Debt(TimeStampedModel):
    """
    tbl_debts — a credit sale.
    Note: in Access this PK was confusingly named sale_id. Here it's just `id`.
    """
    debtor = models.ForeignKey(
        Debtor, on_delete=models.PROTECT, related_name='debts'
    )
    product_spec = models.ForeignKey(
        ProductSpec, on_delete=models.PROTECT, related_name='debts'
    )
    quantity = models.PositiveIntegerField(default=1)
    unit_price = models.DecimalField(max_digits=12, decimal_places=2)
    discount = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    sale_date = models.DateTimeField(default=timezone.now, db_index=True)
    expected_payment_date = models.DateField(null=True, blank=True)

    class Meta:
        db_table = 'credit_debt'
        ordering = ['-sale_date']

    def __str__(self):
        return f"Debt: {self.debtor.name} — TZS {self.amount}"

    @property
    def amount_paid(self):
        return self.debt_returns.aggregate(
            total=Coalesce(Sum('amount'), 0)
        )['total']

    @property
    def balance(self):
        return self.amount - self.amount_paid

    @property
    def is_overdue(self):
        if self.expected_payment_date:
            return self.balance > 0 and self.expected_payment_date < timezone.now().date()
        return False

    def save(self, *args, **kwargs):
        # Auto-calculate amount
        self.amount = (self.quantity * self.unit_price) - self.discount
        super().save(*args, **kwargs)


class DebtReturn(TimeStampedModel):
    """tbl_debt_returns — a payment from a debtor"""
    debt = models.ForeignKey(
        Debt, on_delete=models.PROTECT, related_name='debt_returns'
    )
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    return_date = models.DateTimeField(default=timezone.now)
    payment_method = models.ForeignKey(
        PaymentMethod, on_delete=models.PROTECT
    )
    comment = models.TextField(blank=True)

    class Meta:
        db_table = 'credit_debt_return'
        ordering = ['-return_date']

    def __str__(self):
        return f"Payment of TZS {self.amount} for {self.debt}"
```

### Finance Models (apps/finance/models.py)

```python
# apps/finance/models.py
from django.db import models
from django.utils import timezone
from apps.core.models import TimeStampedModel
from apps.sales.models import PaymentMethod


class ExpenseType(TimeStampedModel):
    """tbl_expense_types"""
    name = models.CharField(max_length=255, unique=True)

    class Meta:
        db_table = 'finance_expense_type'

    def __str__(self):
        return self.name


class ExpenseItem(TimeStampedModel):
    """tbl_expense_items — a specific recurring expense (e.g. 'Monthly Rent')"""
    expense_type = models.ForeignKey(
        ExpenseType, on_delete=models.PROTECT, related_name='items'
    )
    name = models.CharField(max_length=255)
    description = models.TextField(blank=True)
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        db_table = 'finance_expense_item'

    def __str__(self):
        return f"{self.expense_type.name} — {self.name}"

    @property
    def current_rate(self):
        """Get the currently effective rate — equivalent to modRATE logic."""
        return self.rates.filter(
            effective_from__lte=timezone.now().date()
        ).order_by('-effective_from').first()


class ExpenseRate(TimeStampedModel):
    """tbl_expense_rates — rate history for an expense item"""
    expense_item = models.ForeignKey(
        ExpenseItem, on_delete=models.PROTECT, related_name='rates'
    )
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    effective_from = models.DateField()
    change_reason = models.TextField(blank=True)

    class Meta:
        db_table = 'finance_expense_rate'
        ordering = ['-effective_from']

    def __str__(self):
        return f"{self.expense_item.name}: TZS {self.amount} from {self.effective_from}"


class RecurrencePattern(TimeStampedModel):
    """tbl_recurrence_patterns — modExpenses logic in model form"""
    RECURRENCE_CHOICES = [
        ('daily', 'Daily'),
        ('weekly', 'Weekly'),
        ('monthly', 'Monthly'),
        ('quarterly', 'Quarterly'),
        ('annually', 'Annually'),
    ]
    expense_item = models.OneToOneField(
        ExpenseItem, on_delete=models.CASCADE, related_name='recurrence'
    )
    recurrence_type = models.CharField(max_length=50, choices=RECURRENCE_CHOICES)
    frequency_value = models.PositiveIntegerField(default=1)
    specific_day_of_week = models.SmallIntegerField(null=True, blank=True)
    specific_day_of_month = models.SmallIntegerField(null=True, blank=True, default=-1)
    start_date = models.DateField()
    end_date = models.DateField(null=True, blank=True)
    is_active = models.BooleanField(default=True)

    class Meta:
        db_table = 'finance_recurrence_pattern'


class PaymentObligation(TimeStampedModel):
    """tbl_payment_obligations — a bill that is due"""
    OBLIGATION_TYPES = [
        ('expense', 'Expense'),
        ('liability', 'Liability Payment'),
    ]
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('partial', 'Partially Paid'),
        ('paid', 'Paid'),
        ('overdue', 'Overdue'),
        ('cancelled', 'Cancelled'),
    ]
    expense_item = models.ForeignKey(
        ExpenseItem, on_delete=models.PROTECT,
        related_name='obligations', null=True, blank=True
    )
    obligation_type = models.CharField(max_length=50, choices=OBLIGATION_TYPES)
    obligation_date = models.DateField()
    due_date = models.DateField(db_index=True)
    amount_due = models.DecimalField(max_digits=12, decimal_places=2)
    prepayment_applied = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    amount_paid = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    description = models.TextField(blank=True)

    class Meta:
        db_table = 'finance_payment_obligation'
        ordering = ['due_date']

    @property
    def balance(self):
        return self.amount_due - self.prepayment_applied - self.amount_paid

    @property
    def payment_status(self):
        if self.balance <= 0:
            return 'paid'
        if self.amount_paid > 0:
            return 'partial'
        if self.due_date < timezone.now().date():
            return 'overdue'
        return 'pending'
```

---

## 7. Django Admin — Your Free frm_master

This single file gives you a complete, working backoffice interface for all your data. Build this on Day 1.

```python
# apps/catalog/admin.py
from django.contrib import admin
from django.utils.html import format_html
from .models import (
    Category, ProductType, Brand, Unit,
    Spec, SpecValue, Product, ProductSpec
)


class ProductSpecInline(admin.TabularInline):
    """Allows editing product variants inside the Product form."""
    model = ProductSpec
    extra = 1
    fields = ['spec_value', 'default_cost_price', 'default_selling_price',
              'reorder_level', 'current_stock']
    readonly_fields = ['current_stock']


@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = [
        'name', 'product_type', 'brand', 'unit',
        'spec_count', 'stock_status'
    ]
    list_filter = ['product_type__category', 'product_type', 'brand']
    search_fields = ['name', 'barcode']
    inlines = [ProductSpecInline]

    def spec_count(self, obj):
        return obj.specs.count()
    spec_count.short_description = 'Variants'

    def stock_status(self, obj):
        low_stock = obj.specs.filter(
            current_stock__lte=models.F('reorder_level')
        ).count()
        if low_stock:
            return format_html(
                '<span style="color:red;">⚠️ {} low</span>', low_stock
            )
        return format_html('<span style="color:green;">✅ OK</span>')
    stock_status.short_description = 'Stock'


@admin.register(ProductSpec)
class ProductSpecAdmin(admin.ModelAdmin):
    list_display = [
        'product', 'spec_value',
        'default_cost_price', 'default_selling_price',
        'current_stock', 'reorder_level', 'stock_alert'
    ]
    list_filter = ['product__product_type__category', 'product__brand']
    search_fields = ['product__name', 'spec_value__value']
    list_editable = ['default_selling_price', 'reorder_level']

    def stock_alert(self, obj):
        if obj.is_low_stock:
            return format_html('<span style="color:red;">⚠️ LOW</span>')
        return '✅'
    stock_alert.short_description = 'Alert'


# apps/sales/admin.py
from django.contrib import admin
from .models import Sale, SaleReceipt, PaymentMethod


class SaleInline(admin.TabularInline):
    model = Sale
    extra = 0
    fields = ['product_spec', 'quantity', 'unit_price', 'discount', 'amount']
    readonly_fields = ['amount']


@admin.register(SaleReceipt)
class SaleReceiptAdmin(admin.ModelAdmin):
    list_display = ['receipt_number', 'sale_date', 'payment_method', 'total_amount']
    list_filter = ['payment_method', 'sale_date']
    search_fields = ['receipt_number']
    date_hierarchy = 'sale_date'
    inlines = [SaleInline]
    readonly_fields = ['receipt_number']


@admin.register(Sale)
class SaleAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'product_spec', 'quantity', 'unit_price',
        'discount', 'amount', 'payment_method', 'sale_date'
    ]
    list_filter = ['payment_method', 'sale_date']
    search_fields = ['product_spec__product__name']
    date_hierarchy = 'sale_date'
    readonly_fields = ['amount']


# apps/credit/admin.py
from django.contrib import admin
from .models import Debtor, Debt, DebtReturn


class DebtInline(admin.TabularInline):
    model = Debt
    extra = 0
    readonly_fields = ['amount', 'amount_paid', 'balance', 'is_overdue']


class DebtReturnInline(admin.TabularInline):
    model = DebtReturn
    extra = 1


@admin.register(Debtor)
class DebtorAdmin(admin.ModelAdmin):
    list_display = [
        'name', 'phone_1', 'outstanding_balance', 'has_overdue', 'is_active'
    ]
    search_fields = ['name', 'phone_1', 'nida_id']
    list_filter = ['is_active']
    inlines = [DebtInline]


@admin.register(Debt)
class DebtAdmin(admin.ModelAdmin):
    list_display = [
        'debtor', 'product_spec', 'amount', 'amount_paid',
        'balance', 'expected_payment_date', 'is_overdue'
    ]
    list_filter = ['sale_date', 'debtor']
    search_fields = ['debtor__name']
    inlines = [DebtReturnInline]
```

---

## 8. Views, URLs & Business Logic

### Service Layer Pattern (The VBA Modules Replacement)

This is the most important architectural decision. Your VBA modules become Python service classes. They live in `services.py` in each app. They are pure Python — no Django request/response, no templates. Just business logic.

```python
# apps/inventory/services.py
from django.db import transaction
from decimal import Decimal
from apps.catalog.models import ProductSpec


class StockService:
    """
    Equivalent to modStockCRUD in Access VBA.
    All stock operations go through here — never update stock directly.
    """

    @staticmethod
    @transaction.atomic
    def increase_stock(product_spec: ProductSpec, quantity: int) -> ProductSpec:
        """Called after a purchase is confirmed."""
        product_spec.current_stock = models.F('current_stock') + quantity
        product_spec.save(update_fields=['current_stock', 'updated_at'])
        product_spec.refresh_from_db()
        return product_spec

    @staticmethod
    @transaction.atomic
    def decrease_stock(product_spec: ProductSpec, quantity: int) -> ProductSpec:
        """
        Called after a sale. Raises ValueError if insufficient stock.
        """
        # Re-fetch with select_for_update to prevent race conditions
        spec = ProductSpec.objects.select_for_update().get(pk=product_spec.pk)

        if spec.current_stock < quantity:
            raise ValueError(
                f"Insufficient stock for {spec}. "
                f"Available: {spec.current_stock}, Requested: {quantity}"
            )

        spec.current_stock -= quantity
        spec.save(update_fields=['current_stock', 'updated_at'])

        return spec

    @staticmethod
    def get_low_stock_items():
        """Returns all product specs at or below reorder level."""
        from django.db.models import F
        return (
            ProductSpec.objects
            .filter(current_stock__lte=F('reorder_level'))
            .select_related('product', 'product__brand', 'spec_value')
            .order_by('current_stock')
        )

    @staticmethod
    def get_stock_value():
        """Total inventory value at cost price."""
        from django.db.models import F, Sum, ExpressionWrapper, DecimalField
        result = ProductSpec.objects.annotate(
            line_value=ExpressionWrapper(
                F('current_stock') * F('default_cost_price'),
                output_field=DecimalField()
            )
        ).aggregate(total=Sum('line_value'))
        return result['total'] or Decimal('0')


# apps/finance/services.py
from django.utils import timezone
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta  # pip install python-dateutil
from .models import ExpenseItem, PaymentObligation, RecurrencePattern


class ExpenseService:
    """
    Equivalent to modExpenses in Access VBA.
    Generates payment obligations from recurrence patterns.
    """

    @staticmethod
    def generate_obligations_for_month(year: int, month: int) -> list:
        """
        Generate PaymentObligation records for all active recurring expenses
        in the given month. Call this via a management command or cron.
        """
        generated = []
        active_patterns = RecurrencePattern.objects.filter(
            is_active=True,
            expense_item__is_active=True,
        ).select_related('expense_item')

        target_month_start = date(year, month, 1)
        target_month_end = target_month_start + relativedelta(months=1) - timedelta(days=1)

        for pattern in active_patterns:
            due_date = ExpenseService._calculate_due_date(pattern, year, month)
            if due_date is None:
                continue

            # Don't duplicate
            already_exists = PaymentObligation.objects.filter(
                expense_item=pattern.expense_item,
                due_date=due_date,
            ).exists()

            if not already_exists:
                rate = pattern.expense_item.current_rate
                if rate:
                    obligation = PaymentObligation.objects.create(
                        expense_item=pattern.expense_item,
                        obligation_type='expense',
                        obligation_date=timezone.now().date(),
                        due_date=due_date,
                        amount_due=rate.amount,
                        description=f"Auto-generated: {pattern.expense_item.name} — {target_month_start.strftime('%B %Y')}",
                    )
                    generated.append(obligation)

        return generated

    @staticmethod
    def _calculate_due_date(pattern: RecurrencePattern, year: int, month: int):
        if pattern.recurrence_type == 'monthly':
            day = pattern.specific_day_of_month
            if day == -1:
                day = 1
            try:
                return date(year, month, day)
            except ValueError:
                return None
        # Add weekly, quarterly etc. as needed
        return None


# apps/reports/generators/credit.py
from django.db.models import Sum, F, Case, When, DecimalField, Value
from django.db.models.functions import Coalesce
from django.utils import timezone
from apps.credit.models import Debtor, Debt, DebtReturn


class DebtorAgeingGenerator:
    """
    Equivalent to the ageing query in the Access master document.
    Pythonic, testable, no raw SQL needed.
    """

    def generate(self) -> list[dict]:
        today = timezone.now().date()
        results = []

        debtors = Debtor.objects.filter(is_active=True).prefetch_related(
            'debts__debt_returns'
        )

        for debtor in debtors:
            buckets = {'current': 0, 'days_1_30': 0, 'days_31_60': 0, 'days_over_60': 0}

            for debt in debtor.debts.all():
                remaining = debt.balance
                if remaining <= 0:
                    continue

                if debt.expected_payment_date is None:
                    buckets['current'] += remaining
                    continue

                overdue_days = (today - debt.expected_payment_date).days

                if overdue_days <= 0:
                    buckets['current'] += remaining
                elif overdue_days <= 30:
                    buckets['days_1_30'] += remaining
                elif overdue_days <= 60:
                    buckets['days_31_60'] += remaining
                else:
                    buckets['days_over_60'] += remaining

            total = sum(buckets.values())
            if total > 0:
                results.append({
                    'debtor': debtor,
                    'phone': debtor.phone_1,
                    **buckets,
                    'total': total,
                    'is_overdue': buckets['days_1_30'] + buckets['days_31_60'] + buckets['days_over_60'] > 0,
                })

        return sorted(results, key=lambda x: x['days_over_60'], reverse=True)
```

### Dashboard View (apps/dashboard/views.py)

```python
# apps/dashboard/views.py
from django.views.generic import TemplateView
from django.utils import timezone
from django.db.models import Sum
from django.db.models.functions import Coalesce
from apps.sales.models import Sale
from apps.inventory.services import StockService
from apps.credit.models import Debt, DebtReturn


class DashboardView(TemplateView):
    """
    The Django equivalent of frm_master.
    Shows KPIs and navigation. Replaces the Access dashboard.
    """
    template_name = 'dashboard/index.html'

    def get_context_data(self, **kwargs):
        ctx = super().get_context_data(**kwargs)
        today = timezone.now().date()

        # Today's sales
        ctx['today_sales'] = Sale.objects.filter(
            sale_date__date=today
        ).aggregate(total=Coalesce(Sum('amount'), 0))['total']

        # This month's sales
        ctx['month_sales'] = Sale.objects.filter(
            sale_date__year=today.year,
            sale_date__month=today.month,
        ).aggregate(total=Coalesce(Sum('amount'), 0))['total']

        # Low stock count
        ctx['low_stock_count'] = StockService.get_low_stock_items().count()

        # Outstanding debt
        total_debt = Debt.objects.aggregate(t=Coalesce(Sum('amount'), 0))['t']
        total_paid = DebtReturn.objects.aggregate(t=Coalesce(Sum('amount'), 0))['t']
        ctx['outstanding_debt'] = total_debt - total_paid

        # Overdue obligations
        from apps.finance.models import PaymentObligation
        ctx['overdue_obligations'] = PaymentObligation.objects.filter(
            due_date__lt=today, amount_paid__lt=models.F('amount_due')
        ).count()

        return ctx
```

---

## 9. Reports in Django — WeasyPrint + Charts

### PDF Report Generation

```python
# apps/reports/views.py
from django.http import HttpResponse
from django.template.loader import render_to_string
from django.utils import timezone
import weasyprint  # pip install weasyprint


def generate_pdf_report(template_name: str, context: dict, filename: str) -> HttpResponse:
    """
    Universal PDF generator.
    Works for all reports: daily sales, stock levels, debtor ageing, P&L etc.
    """
    html_string = render_to_string(template_name, context)
    html = weasyprint.HTML(string=html_string)
    css = weasyprint.CSS(string='''
        @page {
            size: A4;
            margin: 20mm 15mm;
            @top-center { content: "Kiyabo Duka"; font-size: 10pt; }
            @bottom-right { content: "Page " counter(page) " of " counter(pages); font-size: 9pt; }
        }
        body { font-family: "DejaVu Sans", sans-serif; font-size: 10pt; }
        table { width: 100%; border-collapse: collapse; }
        th { background: #1E3A5F; color: white; padding: 6px; }
        td { padding: 4px 6px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background: #f5f5f5; }
        .total-row { font-weight: bold; border-top: 2px solid #333; }
    ''')
    pdf = html.write_pdf(stylesheets=[css])
    response = HttpResponse(pdf, content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="{filename}"'
    return response


# Daily sales report view
class DailySalesReportView(View):
    def get(self, request):
        from apps.reports.generators.sales import DailySalesGenerator
        date_str = request.GET.get('date', timezone.now().strftime('%Y-%m-%d'))
        report_date = datetime.strptime(date_str, '%Y-%m-%d').date()

        generator = DailySalesGenerator()
        context = generator.generate(report_date)
        context['report_date'] = report_date
        context['generated_at'] = timezone.now()

        if request.GET.get('format') == 'pdf':
            return generate_pdf_report(
                'reports/daily_sales.html',
                context,
                f'daily_sales_{date_str}.pdf'
            )
        return render(request, 'reports/daily_sales.html', context)
```

### Chart Data (using Chart.js via JSON)

```python
# Monthly sales trend — returns JSON for Chart.js
class MonthlySalesTrendView(View):
    def get(self, request):
        from django.db.models.functions import TruncMonth
        year = int(request.GET.get('year', timezone.now().year))

        data = (
            Sale.objects
            .filter(sale_date__year=year)
            .annotate(month=TruncMonth('sale_date'))
            .values('month')
            .annotate(total=Sum('amount'), count=Count('id'))
            .order_by('month')
        )

        return JsonResponse({
            'labels': [d['month'].strftime('%B') for d in data],
            'sales': [float(d['total']) for d in data],
            'transactions': [d['count'] for d in data],
        })
```

---

## 10. Migration Strategy — Access to PostgreSQL

### Step 1: Export from Access

```python
# migration/export_access.py
# Run on the Windows machine with the .accdb file
import pyodbc  # pip install pyodbc
import json
from datetime import datetime, date
from decimal import Decimal


def serialize(obj):
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return float(obj)
    return str(obj)


def export_table(conn, table_name: str, output_path: str):
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}]")
    columns = [col[0] for col in cursor.description]
    rows = []
    for row in cursor.fetchall():
        rows.append(dict(zip(columns, row)))
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(rows, f, default=serialize, indent=2, ensure_ascii=False)
    print(f"Exported {len(rows)} rows from {table_name}")


conn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"Dbq=E:\duka\Kiyabo Duka v0.031.accdb;"
)
conn = pyodbc.connect(conn_str)

TABLES_TO_EXPORT = [
    'tbl_categories', 'tbl_types', 'tbl_brands', 'tbl_units',
    'tbl_specs', 'tbl_spec_values', 'tbl_products', 'tbl_product_specs',
    'tbl_suppliers', 'tbl_purchases', 'tbl_purchase_details',
    'tbl_payment_methods', 'tbl_sales', 'tbl_sales_office_use',
    'tbl_return_inwards', 'tbl_return_outwards',
    'tbl_debtors', 'tbl_debts', 'tbl_debt_returns',
    'tbl_expense_types', 'tbl_expense_items', 'tbl_expense_rates',
    'tbl_recurrence_patterns', 'tbl_payment_obligations',
    'tbl_payments', 'tbl_prepayments', 'tbl_payment_allocations',
    'tbl_drawing_categories', 'tbl_drawings',
    'tbl_asset_categories', 'tbl_asset_types', 'tbl_assets',
    'tbl_liability_categories', 'tbl_liability_types',
    'tbl_liability_items', 'tbl_liability_payment_details',
]

for table in TABLES_TO_EXPORT:
    export_table(conn, table, f'migration/data/{table}.json')

conn.close()
print("Export complete.")
```

### Step 2: Django Management Command to Import

```python
# apps/core/management/commands/import_access_data.py
import json
from pathlib import Path
from django.core.management.base import BaseCommand
from django.db import transaction


class Command(BaseCommand):
    help = 'Import data exported from Access database'

    DATA_DIR = Path('migration/data')

    def handle(self, *args, **options):
        self.stdout.write('Starting Access data import...')

        with transaction.atomic():
            self.import_catalog()
            self.import_suppliers()
            self.import_sales()
            self.import_credit()
            self.import_finance()
            self.import_assets()

        self.stdout.write(self.style.SUCCESS('Import complete!'))

    def load(self, table_name: str) -> list:
        path = self.DATA_DIR / f'{table_name}.json'
        if not path.exists():
            self.stdout.write(f'  WARNING: {table_name}.json not found, skipping')
            return []
        with open(path, encoding='utf-8') as f:
            return json.load(f)

    def import_catalog(self):
        from apps.catalog.models import (
            Category, ProductType, Brand, Unit, Spec, SpecValue, Product, ProductSpec
        )
        self.stdout.write('  Importing catalog...')

        # Categories
        cat_map = {}  # old_id → new_id
        for row in self.load('tbl_categories'):
            obj, _ = Category.objects.get_or_create(name=row['category_name'])
            cat_map[row['category_id']] = obj.id

        # Types
        type_map = {}
        for row in self.load('tbl_types'):
            cat = Category.objects.get(pk=cat_map.get(row['category_id']))
            obj, _ = ProductType.objects.get_or_create(
                category=cat, name=row['type_name']
            )
            type_map[row['type_id']] = obj.id

        # Brands
        brand_map = {}
        for row in self.load('tbl_brands'):
            obj, _ = Brand.objects.get_or_create(name=row['brand_name'])
            brand_map[row['brand_id']] = obj.id

        # Units
        unit_map = {}
        for row in self.load('tbl_units'):
            obj, _ = Unit.objects.get_or_create(
                name=row['unit_name'],
                defaults={'abbreviation': row.get('unit_abbr', '')}
            )
            unit_map[row['unit_id']] = obj.id

        # Specs
        spec_map = {}
        for row in self.load('tbl_specs'):
            obj, _ = Spec.objects.get_or_create(name=row['spec_name'])
            spec_map[row['spec_id']] = obj.id

        # SpecValues
        sv_map = {}
        for row in self.load('tbl_spec_values'):
            spec = Spec.objects.get(pk=spec_map.get(row['spec_id']))
            obj, _ = SpecValue.objects.get_or_create(
                spec=spec, value=row['spec_value']
            )
            sv_map[row['spec_value_id']] = obj.id

        # Products
        prod_map = {}
        for row in self.load('tbl_products'):
            ptype = ProductType.objects.get(pk=type_map.get(row['type_id']))
            brand = Brand.objects.get(pk=brand_map[row['brand_id']]) if row.get('brand_id') else None
            unit = Unit.objects.get(pk=unit_map[row['unit_id']]) if row.get('unit_id') else None
            obj, _ = Product.objects.get_or_create(
                name=row['product_name'],
                product_type=ptype,
                defaults={'brand': brand, 'unit': unit}
            )
            prod_map[row['product_id']] = obj.id

        # ProductSpecs
        ps_map = {}
        for row in self.load('tbl_product_specs'):
            prod = Product.objects.get(pk=prod_map.get(row['product_id']))
            sv = SpecValue.objects.get(pk=sv_map.get(row['spec_value_id']))
            obj, _ = ProductSpec.objects.get_or_create(
                product=prod, spec_value=sv,
                defaults={
                    'default_cost_price': row.get('default_cost_price'),
                    'default_selling_price': row.get('default_selling_price'),
                    'reorder_level': row.get('reorder_level', 5),
                    'current_stock': row.get('current_stock', 0),
                }
            )
            ps_map[row['product_spec_id']] = obj.id

        self.stdout.write(
            f'  ✅ Catalog: {Product.objects.count()} products, '
            f'{ProductSpec.objects.count()} specs'
        )
        return ps_map  # Return for use by other importers
```

---

## 11. Deployment — Production-Ready Setup

### Technology Stack

```
Ubuntu 22.04 / 24.04 LTS
├── PostgreSQL 16          ← database
├── Redis 7                ← cache + Celery broker
├── Gunicorn               ← WSGI server (Django)
├── Nginx                  ← reverse proxy + static files
├── Celery                 ← background tasks (expense generation etc.)
├── Celery Beat            ← scheduled tasks (monthly obligation generation)
└── Supervisor             ← process management
```

### requirements.txt

```
# Core
Django==5.1.*
psycopg2-binary==2.9.*
python-decouple==3.8.*
Pillow==10.*

# Date handling (for recurrence calculations)
python-dateutil==2.9.*

# PDF reports
weasyprint==62.*

# Background tasks
celery==5.4.*
redis==5.*
django-celery-beat==2.7.*

# Charts (optional — use Chart.js via CDN instead)
# matplotlib  ← only if you need server-side chart images

# Development tools
django-debug-toolbar==4.*
django-extensions==3.*

# Production
gunicorn==22.*
whitenoise==6.*        # serve static files without Nginx for simple deploys
sentry-sdk==2.*        # error tracking
```

### config/settings/production.py

```python
from .base import *
from decouple import config, Csv

DEBUG = False
SECRET_KEY = config('SECRET_KEY')
ALLOWED_HOSTS = config('ALLOWED_HOSTS', cast=Csv())

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': config('DB_NAME'),
        'USER': config('DB_USER'),
        'PASSWORD': config('DB_PASSWORD'),
        'HOST': config('DB_HOST', default='localhost'),
        'PORT': config('DB_PORT', default='5432'),
        'CONN_MAX_AGE': 60,  # Connection pooling
    }
}

# Cache
CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.redis.RedisCache',
        'LOCATION': config('REDIS_URL', default='redis://localhost:6379/1'),
    }
}

# Static & Media
STATIC_ROOT = BASE_DIR / 'staticfiles'
MEDIA_ROOT = BASE_DIR / 'media'
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

# Security
SECURE_SSL_REDIRECT = True
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True
SECURE_HSTS_SECONDS = 31536000

# Celery
CELERY_BROKER_URL = config('REDIS_URL', default='redis://localhost:6379/0')
CELERY_RESULT_BACKEND = config('REDIS_URL', default='redis://localhost:6379/0')
```

### Celery for Background Tasks

This replaces your VBA module scheduled tasks:

```python
# config/celery.py
import os
from celery import Celery
from celery.schedules import crontab

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings.production')
app = Celery('kiyabo_duka')
app.config_from_object('django.conf:settings', namespace='CELERY')
app.autodiscover_tasks()

app.conf.beat_schedule = {
    # Generate next month's expense obligations on the 25th of each month
    'generate-monthly-obligations': {
        'task': 'apps.finance.tasks.generate_next_month_obligations',
        'schedule': crontab(day_of_month=25, hour=8, minute=0),
    },
    # Daily stock level check and alert at 8am
    'daily-stock-check': {
        'task': 'apps.inventory.tasks.check_low_stock',
        'schedule': crontab(hour=8, minute=0),
    },
}


# apps/finance/tasks.py
from celery import shared_task
from django.utils import timezone


@shared_task
def generate_next_month_obligations():
    """Replaces the manual 'generate obligations' process in Access."""
    from apps.finance.services import ExpenseService
    today = timezone.now().date()
    # Generate for next month
    next_month = today.replace(day=1) + relativedelta(months=1)
    obligations = ExpenseService.generate_obligations_for_month(
        next_month.year, next_month.month
    )
    return f"Generated {len(obligations)} obligations for {next_month.strftime('%B %Y')}"
```

### Where to Host

| Option | Cost/Month | Best For |
|---|---|---|
| **Railway.app** | $5–20 | Easiest Django deployment, free tier available |
| **Render.com** | $7–25 | Good DX, PostgreSQL included |
| **DigitalOcean Droplet** | $6–12 | Full control, cheapest for VPS |
| **PythonAnywhere** | $5–12 | Python-specific, very easy, good for beginners |
| **Local Ubuntu server** | Power cost only | If internet isn't needed, cheapest possible |

**For a single-shop system in Dar es Salaam with no need for remote access:**  
A DigitalOcean $6/month droplet (1GB RAM) with Ubuntu + PostgreSQL + Nginx is more than sufficient and fully within your control.

**If you need mobile/tablet access from the shop floor:**  
Railway or Render — deploy in 20 minutes, HTTPS automatic, no server management.

---

## 12. FastAPI + Next.js Deep Dive

If you do eventually go this route (mobile app, multiple branches, real-time dashboard), here is the honest implementation picture.

### What the Stack Looks Like

```
kiyabo_duka_api/          ← FastAPI backend
├── main.py               ← FastAPI app, router registration
├── core/
│   ├── database.py       ← SQLAlchemy engine + session
│   ├── config.py         ← Pydantic Settings
│   └── security.py       ← JWT token handling
├── models/               ← SQLAlchemy ORM models (same tables, different syntax)
├── schemas/              ← Pydantic request/response models
├── routers/              ← API route handlers
│   ├── sales.py
│   ├── products.py
│   ├── debtors.py
│   └── reports.py
├── services/             ← Business logic (identical to Django services)
└── alembic/              ← Database migrations (manual, unlike Django)

kiyabo_duka_web/          ← Next.js frontend (separate repo)
├── app/                  ← Next.js 14 App Router
│   ├── dashboard/
│   ├── sales/
│   ├── products/
│   └── reports/
├── components/           ← React components
├── lib/
│   └── api.ts            ← API client (fetch wrappers)
└── types/                ← TypeScript type definitions
```

### FastAPI Example — Sales Endpoint

```python
# routers/sales.py
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from datetime import date, datetime
from typing import Optional
from core.database import get_db
from schemas.sales import SaleCreate, SaleResponse, DailySalesSummary
from services.sales import SalesService
from services.stock import StockService

router = APIRouter(prefix="/api/v1/sales", tags=["Sales"])


@router.post("/", response_model=SaleResponse, status_code=201)
async def create_sale(sale_data: SaleCreate, db: Session = Depends(get_db)):
    """
    Record a new sale.
    Automatically decrements stock.
    """
    # Check stock first
    product_spec = db.get(ProductSpec, sale_data.product_spec_id)
    if not product_spec:
        raise HTTPException(status_code=404, detail="Product spec not found")

    if product_spec.current_stock < sale_data.quantity:
        raise HTTPException(
            status_code=400,
            detail=f"Insufficient stock. Available: {product_spec.current_stock}"
        )

    service = SalesService(db)
    sale = service.create_sale(sale_data)
    return sale


@router.get("/daily-summary", response_model=DailySalesSummary)
async def daily_summary(
    report_date: date = Query(default_factory=date.today),
    db: Session = Depends(get_db)
):
    """Returns the daily sales summary — used by the dashboard and report."""
    service = SalesService(db)
    return service.get_daily_summary(report_date)


@router.get("/monthly-trend")
async def monthly_trend(
    year: int = Query(default_factory=lambda: datetime.now().year),
    db: Session = Depends(get_db)
):
    """Returns monthly sales data for charting."""
    service = SalesService(db)
    return service.get_monthly_trend(year)
```

### What You Gain Over Django in This Stack

1. **Auto-generated API docs** at `/docs` (Swagger UI) — useful if you ever integrate M-Pesa, other systems
2. **True async** — if you have high concurrent users (multiple cashiers simultaneously), FastAPI handles this more naturally
3. **WebSockets** — real-time dashboard updates (`sale just recorded: TZS 5,000`) are trivial in FastAPI
4. **Typed API contracts** — Pydantic schemas document your API and validate input at the boundary
5. **React frontend** — if you eventually want a proper mobile-responsive web app or React Native mobile app

### What You Lose

- Django Admin (you build every CRUD screen from scratch in React)
- Django ORM convenience (SQLAlchemy is more powerful but more verbose)
- Built-in auth (you wire JWT manually)
- Built-in migrations (Alembic is manual)
- Single deployment (now two deployments: API + frontend)
- You write JavaScript/TypeScript (React, Next.js) — if you want to stay purely Pythonic, this is the wrong choice

---

## 13. The Decision Framework

### Use This Flowchart

```
Is your Python comfort high and JavaScript comfort low?
    → YES → Django. Full stop.

Will this system run on ONE computer / LOCAL NETWORK only?
    → YES → Django. Access even. No need for web stack.

Do you need mobile access (phone/tablet in the shop)?
    → YES → Django (browser-based) works fine on mobile
             FastAPI+Next.js only if you need a native app

Do you plan to build a mobile app (React Native)?
    → YES → FastAPI+Next.js is a better foundation
    → NO  → Django

Will there be multiple shops / branches?
    → YES → Django with DRF API is still simpler. FastAPI if 5+ branches.

Is real-time dashboard (sales updating live) critical?
    → YES → Django Channels OR FastAPI. Both work.
    → NO  → Django is simpler.

Are you the only developer?
    → YES → Django. Two codebases = 2× maintenance.

CONCLUSION FOR KIYABO DUKA TODAY:
→ Django monolith
→ Add DRF (Django REST Framework) when you need an API
→ Consider FastAPI only when you have a compelling async/mobile reason
```

### The Hybrid Approach (Best Long-Term Architecture)

```
Core system: Django monolith (forms, reports, admin, CRUD)
    + Django REST Framework for any API needs

Optional additions (when needed):
    FastAPI microservice for webhooks (M-Pesa, SMS, etc.)
    React/Next.js for specific high-UX screens (POS terminal)
    Django Channels for real-time dashboard
```

This gives you: Django productivity now + escape hatch to modern stack when justified.

---

## 14. Complete Django Boilerplate

### One-Command Project Setup

```bash
# Create and activate virtualenv
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows

# Install Django and core deps
pip install django psycopg2-binary python-decouple weasyprint python-dateutil

# Create project with config package
django-admin startproject config .

# Create apps
mkdir apps
python manage.py startapp core apps/core
python manage.py startapp catalog apps/catalog
python manage.py startapp inventory apps/inventory
python manage.py startapp sales apps/sales
python manage.py startapp credit apps/credit
python manage.py startapp finance apps/finance
python manage.py startapp assets apps/assets
python manage.py startapp reports apps/reports
python manage.py startapp dashboard apps/dashboard

# Create initial migrations
python manage.py makemigrations
python manage.py migrate

# Create superuser (your admin account)
python manage.py createsuperuser

# Run dev server
python manage.py runserver
```

### config/settings/base.py

```python
from pathlib import Path
from decouple import config

BASE_DIR = Path(__file__).resolve().parent.parent.parent

SECRET_KEY = config('SECRET_KEY', default='dev-insecure-key-change-in-production')
DEBUG = config('DEBUG', default=True, cast=bool)
ALLOWED_HOSTS = ['*']

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    # Your apps
    'apps.core',
    'apps.catalog',
    'apps.inventory',
    'apps.sales',
    'apps.credit',
    'apps.finance',
    'apps.assets',
    'apps.reports',
    'apps.dashboard',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'config.urls'

TEMPLATES = [{
    'BACKEND': 'django.template.backends.django.DjangoTemplates',
    'DIRS': [BASE_DIR / 'templates'],
    'APP_DIRS': True,
    'OPTIONS': {
        'context_processors': [
            'django.template.context_processors.debug',
            'django.template.context_processors.request',
            'django.contrib.auth.context_processors.auth',
            'django.contrib.messages.context_processors.messages',
        ],
    },
}]

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': config('DB_NAME', default='kiyabo_duka'),
        'USER': config('DB_USER', default='postgres'),
        'PASSWORD': config('DB_PASSWORD', default=''),
        'HOST': config('DB_HOST', default='localhost'),
        'PORT': config('DB_PORT', default='5432'),
    }
}

# Currency / locale
LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Africa/Dar_es_Salaam'  # Tanzania timezone
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Django Admin customization
ADMIN_SITE_HEADER = 'Kiyabo Duka Administration'
ADMIN_SITE_TITLE = 'Kiyabo Duka'
ADMIN_INDEX_TITLE = 'Shop Management'
```

### config/urls.py

```python
from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

# Customize admin header
admin.site.site_header = 'Kiyabo Duka'
admin.site.site_title = 'Kiyabo Duka Admin'
admin.site.index_title = 'Shop Management Dashboard'

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('apps.dashboard.urls')),
    path('catalog/', include('apps.catalog.urls')),
    path('inventory/', include('apps.inventory.urls')),
    path('sales/', include('apps.sales.urls')),
    path('credit/', include('apps.credit.urls')),
    path('finance/', include('apps.finance.urls')),
    path('assets/', include('apps.assets.urls')),
    path('reports/', include('apps.reports.urls')),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
```

---

## Summary — The Three-Line Decision

**You are Pythonic. Your data model is solid. Your system is single-shop.**

1. **Finish Access first** (12 days) — it gives you working software and clarifies all business logic
2. **Build Django in parallel** starting Month 2, using this document as your blueprint
3. **Migrate when Django is feature-complete** — data migration takes one afternoon with the export/import scripts above

Django is not a compromise. Django with PostgreSQL + WeasyPrint + Celery is genuinely production-grade software that runs real businesses at scale. Your Kiyabo Duka system will be cleaner, faster, more maintainable, and more extensible in Django than it ever could be in Access — and you will be writing pure Python the whole time.

FastAPI + Next.js is a great stack. It is just not the right stack for a solo Pythonic developer building a shop management system today. Save it for when you have a mobile app or a second branch that justifies the frontend investment.

---

*Document: Kiyabo Duka Migration Analysis*  
*Version: 1.0 — May 2026*  
*Update this document as decisions are made and architecture evolves*
