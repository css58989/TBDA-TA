"""Catalog of merged-restaurant-info attributes: meaning, expected type, nesting."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class AttributeSpec:
    path: str
    meaning: str
    expected_type: str
    section: str
    parent: Optional[str] = None
    nested_level: int = 0  # 0 top-level, 1 child, 2 grandchild


SECTIONS: list[str] = [
    "1. Restaurant Identity",
    "2. Branch Identity",
    "3. Location",
    "4. Cuisine / Business Type",
    "5. Ratings",
    "6. Delivery & Ordering",
    "7. Status",
    "8. Promotions & Visibility",
    "9. Payment",
    "10. Customer / Delivery Experience",
    "11. Restaurant Services",
    "12. Images & Content",
    "13. VAT / Pricing Configuration",
    "14. Platform / Classification",
    "15. Dates / Lifecycle",
    "16. URLs",
]


ATTRIBUTES: list[AttributeSpec] = [
    # 1. Restaurant Identity
    AttributeSpec("id", "Unique restaurant brand/entity identifier", "int", SECTIONS[0]),
    AttributeSpec("name", "Restaurant display name", "string", SECTIONS[0]),
    AttributeSpec("restaurantId", "Restaurant ID (often same as id)", "int", SECTIONS[0]),
    AttributeSpec("restaurantSlug", "URL-friendly restaurant slug", "string", SECTIONS[0]),
    # 2. Branch Identity
    AttributeSpec("branchId", "Unique branch/outlet identifier", "int", SECTIONS[1]),
    AttributeSpec("branchName", "Branch display name (often includes area)", "string", SECTIONS[1]),
    AttributeSpec("branchSlug", "URL-friendly branch slug", "string", SECTIONS[1]),
    AttributeSpec("branchUrl", "Relative branch page URL path", "string", SECTIONS[1]),
    # 3. Location
    AttributeSpec("areaName", "Human-readable delivery area name", "string", SECTIONS[2]),
    AttributeSpec("areaId", "Delivery area identifier", "int|null", SECTIONS[2]),
    AttributeSpec("deliveryAreaId", "Specific delivery-area ID", "int|null", SECTIONS[2]),
    AttributeSpec("shopArea", "Numeric shop/area code used by platform", "int", SECTIONS[2]),
    AttributeSpec("shopCity", "Numeric city code used by platform", "int", SECTIONS[2]),
    AttributeSpec("latitude", "Branch latitude (often stored as string)", "string|number", SECTIONS[2]),
    AttributeSpec("longitude", "Branch longitude (often stored as string)", "string|number", SECTIONS[2]),
    # 4. Cuisine / Business Type
    AttributeSpec("cuisineString", "Comma-separated cuisine labels", "string", SECTIONS[3]),
    AttributeSpec("cuisines", "Structured list of cuisine objects", "array<object>", SECTIONS[3]),
    AttributeSpec(
        "cuisines.id",
        "Cuisine identifier",
        "int",
        SECTIONS[3],
        parent="cuisines",
        nested_level=1,
    ),
    AttributeSpec(
        "cuisines.name",
        "Cuisine display name",
        "string",
        SECTIONS[3],
        parent="cuisines",
        nested_level=1,
    ),
    AttributeSpec(
        "cuisines.slug",
        "Cuisine URL slug",
        "string",
        SECTIONS[3],
        parent="cuisines",
        nested_level=1,
    ),
    AttributeSpec("verticalType", "Business vertical code (food/grocery/etc.)", "int", SECTIONS[3]),
    AttributeSpec("isGrocery", "Whether the shop is a grocery vertical", "bool", SECTIONS[3]),
    AttributeSpec("shopType", "Shop type code (e.g. TGO, 9C)", "string", SECTIONS[3]),
    AttributeSpec("isDarkstore", "Whether the shop is a darkstore", "bool", SECTIONS[3]),
    AttributeSpec("isCokeRestaurant", "Coca-Cola partnership / Coke flag", "bool", SECTIONS[3]),
    # 5. Ratings
    AttributeSpec("rate", "Average customer rating score", "float", SECTIONS[4]),
    AttributeSpec("totalReviews", "Total number of written reviews", "int", SECTIONS[4]),
    AttributeSpec("totalRatings", "Total number of ratings", "int", SECTIONS[4]),
    # 6. Delivery & Ordering
    AttributeSpec("deliveryFee", "Delivery fee amount (often string)", "string|number", SECTIONS[5]),
    AttributeSpec("minimumOrderAmount", "Minimum order value required", "number", SECTIONS[5]),
    AttributeSpec("avgDeliveryTime", "Human-readable average delivery time text", "string", SECTIONS[5]),
    AttributeSpec("deliveryTime", "Average delivery time in minutes", "int", SECTIONS[5]),
    AttributeSpec("deliveryChargesType", "Delivery-charge calculation type code", "int", SECTIONS[5]),
    AttributeSpec("deliverySchedule", "Current schedule state (e.g. Open)", "string", SECTIONS[5]),
    AttributeSpec("preOrder", "Whether pre-order is available", "bool", SECTIONS[5]),
    AttributeSpec("isTalabatGO", "Whether branch is on Talabat GO", "bool", SECTIONS[5]),
    # 7. Status
    AttributeSpec("statusCode", "Numeric operational status code", "int", SECTIONS[6]),
    AttributeSpec("status", "Status value (often stringified code)", "string|int", SECTIONS[6]),
    AttributeSpec("isNew", "Whether restaurant/branch is marked new", "bool", SECTIONS[6]),
    AttributeSpec("IsMigratedToDh", "Migrated to Delivery Hero stack", "bool", SECTIONS[6]),
    # 8. Promotions & Visibility
    AttributeSpec("promotionText", "Promotion banner/text", "string", SECTIONS[7]),
    AttributeSpec("discountText", "Discount label/text", "string", SECTIONS[7]),
    AttributeSpec("isShopSponcered", "Whether listing is sponsored (source spelling)", "bool", SECTIONS[7]),
    AttributeSpec("Sponsored", "Sponsored placement metadata object", "object", SECTIONS[7]),
    AttributeSpec(
        "Sponsored.category",
        "Sponsorship/placement category",
        "string",
        SECTIONS[7],
        parent="Sponsored",
        nested_level=1,
    ),
    AttributeSpec(
        "Sponsored.type",
        "Sponsorship type (e.g. cpc)",
        "string",
        SECTIONS[7],
        parent="Sponsored",
        nested_level=1,
    ),
    AttributeSpec(
        "Sponsored.token",
        "Opaque sponsored-ad token",
        "string",
        SECTIONS[7],
        parent="Sponsored",
        nested_level=1,
    ),
    AttributeSpec("shopPosition", "Position/rank in listing results", "int", SECTIONS[7]),
    # 9. Payment
    AttributeSpec("acceptCreditCard", "Accepts credit cards", "bool", SECTIONS[8]),
    AttributeSpec("acceptDebitCard", "Accepts debit cards", "bool", SECTIONS[8]),
    AttributeSpec("acceptCash", "Accepts cash on delivery", "bool", SECTIONS[8]),
    AttributeSpec(
        "availablePaymentMethods",
        "List of accepted payment method descriptors",
        "array<object>",
        SECTIONS[8],
    ),
    AttributeSpec(
        "availablePaymentMethods.logo",
        "Payment method logo filename",
        "string",
        SECTIONS[8],
        parent="availablePaymentMethods",
        nested_level=1,
    ),
    AttributeSpec(
        "availablePaymentMethods.text",
        "Payment method label / i18n key",
        "string",
        SECTIONS[8],
        parent="availablePaymentMethods",
        nested_level=1,
    ),
    AttributeSpec(
        "availablePaymentMethods.dimension",
        "Logo dimension object",
        "object",
        SECTIONS[8],
        parent="availablePaymentMethods",
        nested_level=1,
    ),
    AttributeSpec(
        "availablePaymentMethods.dimension.width",
        "Logo width in pixels",
        "int",
        SECTIONS[8],
        parent="availablePaymentMethods.dimension",
        nested_level=2,
    ),
    AttributeSpec(
        "availablePaymentMethods.dimension.height",
        "Logo height in pixels",
        "int",
        SECTIONS[8],
        parent="availablePaymentMethods.dimension",
        nested_level=2,
    ),
    # 10. Customer / Delivery Experience
    AttributeSpec("contactlessDelivery", "Contactless delivery supported", "bool", SECTIONS[9]),
    AttributeSpec("IsProvideTracking", "Live delivery tracking available", "bool", SECTIONS[9]),
    AttributeSpec("isProvideOrderStatus", "Order-status updates provided", "bool", SECTIONS[9]),
    # 11. Restaurant Services
    AttributeSpec("isCateringAvailable", "Catering service available", "bool", SECTIONS[10]),
    AttributeSpec("grlRequired", "GRL (geo/location) requirement flag", "bool", SECTIONS[10]),
    # 12. Images & Content
    AttributeSpec("heroImage", "Cover/hero image URL", "string", SECTIONS[11]),
    AttributeSpec("logo", "Restaurant logo image URL", "string", SECTIONS[11]),
    AttributeSpec("summary", "Short restaurant description/summary", "string", SECTIONS[11]),
    AttributeSpec("altTalabattxt", "Alternate Talabat delivery text", "string", SECTIONS[11]),
    AttributeSpec("altMunicipaltxt", "Alternate municipal delivery text", "string", SECTIONS[11]),
    AttributeSpec("altTouristtxt", "Alternate tourist delivery text", "string", SECTIONS[11]),
    AttributeSpec(
        "alternativeDeliveryText",
        "Alternative delivery/service charge text",
        "string",
        SECTIONS[11],
    ),
    AttributeSpec("branchLearnMoreLink", "Learn-more link for the branch", "string", SECTIONS[11]),
    # 13. VAT / Pricing Configuration
    AttributeSpec("isVatInclusive", "Prices include VAT", "bool", SECTIONS[12]),
    # 14. Platform / Classification
    AttributeSpec("filtersIds", "Filter/tag IDs applied to the shop", "array<int>", SECTIONS[13]),
    AttributeSpec("view", "Listing view/mode code", "int", SECTIONS[13]),
    AttributeSpec("page", "Source listing page number", "int", SECTIONS[13]),
    # 15. Dates / Lifecycle
    AttributeSpec("createdAt", "Creation/lifecycle timestamp-like value", "int", SECTIONS[14]),
    # 16. URLs
    AttributeSpec("branchUrl", "Relative branch URL path", "string", SECTIONS[15]),
    AttributeSpec("menuUrl", "Relative menu URL path", "string", SECTIONS[15]),
]


TOP_LEVEL_PATHS = [a.path for a in ATTRIBUTES if a.nested_level == 0 and "." not in a.path]
# branchUrl appears in sections 2 and 16 in the source list; keep unique for stats
UNIQUE_ATTRIBUTES: list[AttributeSpec] = []
_seen: set[str] = set()
for _attr in ATTRIBUTES:
    if _attr.path in _seen:
        continue
    _seen.add(_attr.path)
    UNIQUE_ATTRIBUTES.append(_attr)
