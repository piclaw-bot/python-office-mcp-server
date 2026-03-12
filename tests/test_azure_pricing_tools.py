"""Tests for azure_pricing_tools.py."""

import pytest

from tools.azure_pricing_tools import AzurePricingTools


@pytest.fixture(autouse=True)
def reset_price_cache():
    """Reset shared price cache before each test."""
    AzurePricingTools._price_cache = {}
    yield
    AzurePricingTools._price_cache = {}


def _cache_key(service: str, region: str, currency: str) -> str:
    return f"{service}_{region}_{currency}".lower().replace(" ", "-")


def test_service_alias_resolves_to_cached_name():
    """Should resolve common Azure-prefixed names to cached service names."""
    tools = AzurePricingTools()
    service = "API Management"
    region = "westeurope"
    currency = "USD"

    AzurePricingTools._price_cache[_cache_key(service, region, currency)] = [
        {
            "serviceName": service,
            "productName": "API Management",
            "skuName": "Developer",
            "armSkuName": "Developer",
            "meterName": "Consumption",
            "retailPrice": 1.23,
            "unitOfMeasure": "1 Hour",
            "type": "Consumption",
        }
    ]

    result = tools.tool_azure_query_prices(
        service="Azure API Management",
        region=region,
        currency=currency,
    )

    assert result.get("service") == service
    assert result.get("service_requested") == "Azure API Management"
    assert result.get("count") == 1


def test_fabric_tier_calculation_uses_cu_mapping():
    """Should translate Fabric F-tier aliases into CU-hour calculations."""
    tools = AzurePricingTools()
    service = "Microsoft Fabric"
    region = "westeurope"
    currency = "USD"

    AzurePricingTools._price_cache[_cache_key(service, region, currency)] = [
        {
            "serviceName": service,
            "productName": "Fabric Capacity",
            "skuName": "Fabric Capacity",
            "armSkuName": "Fabric_Capacity_CU_Hour",
            "meterName": "CU Hour",
            "retailPrice": 0.18,
            "unitOfMeasure": "1 CU-Hour",
            "type": "Consumption",
        }
    ]

    result = tools.tool_azure_calculate_cost(
        service="Azure Fabric",
        region=region,
        currency=currency,
        sku_match="F64",
        quantity=1,
        hours_per_month=1,
    )

    assert result.get("fabric_tier") == "F64"
    assert result.get("fabric_cu") == 64
    assert result.get("effective_quantity") == 64
    assert result.get("monthly_cost") == pytest.approx(0.18 * 64, rel=1e-6)


def test_list_services_from_cache_only():
    """Should list services discovered from cached items."""
    tools = AzurePricingTools()
    region = "westeurope"
    currency = "USD"

    AzurePricingTools._price_cache[_cache_key("Service Bus", region, currency)] = [
        {"serviceName": "Service Bus"}
    ]
    AzurePricingTools._price_cache[_cache_key("Key Vault", region, currency)] = [
        {"serviceName": "Key Vault"}
    ]

    result = tools.tool_azure_list_services(
        region=region,
        currency=currency,
        from_cache_only=True,
    )

    assert result.get("source") == "cache"
    assert set(result.get("services", [])) >= {"Service Bus", "Key Vault"}
