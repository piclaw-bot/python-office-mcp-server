#!/usr/bin/env python3
"""
azure_pricing_tools.py - MCP tools for Azure Retail Prices API.

Provides tools to fetch, cache, and query Azure pricing data for
creating and validating cost estimates.

API Documentation: https://learn.microsoft.com/en-us/rest/api/cost-management/retail-prices/azure-retail-prices
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote

# Check for required libraries
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# API configuration
AZURE_PRICES_API = "https://prices.azure.com/api/retail/prices"
DEFAULT_API_VERSION = "2023-01-01-preview"

# Cache configuration
CACHE_DIR = Path(__file__).parent.parent / "cache" / "azure-pricing"
CACHE_TTL_HOURS = 24  # Cache validity in hours

FABRIC_CAPACITY_TIERS = {
    "F2": 2,
    "F4": 4,
    "F8": 8,
    "F16": 16,
    "F32": 32,
    "F64": 64,
}

SERVICE_ALIAS_MAP = {
    "azure fabric": "Microsoft Fabric",
    "azure api management": "API Management",
    "azure service bus": "Service Bus",
    "azure key vault": "Key Vault",
    "azure functions": "Azure Functions",
    "azure openai": "Azure OpenAI",
}


def _get_cache_path(cache_key: str) -> Path:
    """Get the cache file path for a given key."""
    # Sanitize cache key for filesystem
    safe_key = "".join(c if c.isalnum() or c in "-_" else "_" for c in cache_key)
    return CACHE_DIR / f"{safe_key}.json"


def _is_cache_valid(cache_path: Path, ttl_hours: int = CACHE_TTL_HOURS) -> bool:
    """Check if cache file exists and is not expired."""
    if not cache_path.exists():
        return False

    mtime = cache_path.stat().st_mtime
    age_hours = (time.time() - mtime) / 3600
    return age_hours < ttl_hours


def _load_cache(cache_path: Path) -> dict | None:
    """Load data from cache file."""
    try:
        with open(cache_path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def _save_cache(cache_path: Path, data: dict) -> None:
    """Save data to cache file."""
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def _normalize_name(name: str) -> str:
    """Normalize a service name for comparison."""
    return " ".join(name.lower().split())


def _extract_cache_service_names(cache_items: list[dict]) -> set[str]:
    """Extract distinct service names from cached pricing items."""
    services = set()
    for item in cache_items:
        service_name = item.get("serviceName")
        if service_name:
            services.add(service_name)
    return services


class AzurePricingTools:
    """MCP tool mixin for Azure pricing operations.

    Provides tools to:
    - Fetch and cache Azure retail prices by service/region
    - Query cached pricing data in memory
    - Validate estimates against current prices
    - Build new estimates from resource specifications

    Required packages:
    - requests: HTTP requests (pip install requests)
    """

    # In-memory price cache for fast queries
    _price_cache: dict[str, list[dict]] = {}

    def _resolve_service_name(
        self,
        service: str,
        region: str,
        currency: str,
        allow_fetch: bool = False,
    ) -> tuple[str, list[str]]:
        """Resolve user-friendly service names to API service names."""
        if not service:
            return service, []

        normalized = _normalize_name(service)
        if normalized in SERVICE_ALIAS_MAP:
            return SERVICE_ALIAS_MAP[normalized], [SERVICE_ALIAS_MAP[normalized]]

        candidate_services = set()

        # Collect from memory cache
        for key, items in self._price_cache.items():
            if key.endswith(f"_{region}_{currency}".lower().replace(" ", "-")):
                candidate_services.update(_extract_cache_service_names(items))

        # Collect from disk cache for this region/currency
        if CACHE_DIR.exists():
            for cache_file in CACHE_DIR.glob("*.json"):
                cached = _load_cache(cache_file)
                if not cached:
                    continue
                if cached.get("region") != region or cached.get("currency") != currency:
                    continue
                candidate_services.update(_extract_cache_service_names(cached.get("items", [])))

        if not candidate_services and allow_fetch:
            listing = self.tool_azure_list_services(
                region=region,
                currency=currency,
                from_cache_only=False,
            )
            candidate_services.update(listing.get("services", []))

        if not candidate_services:
            return service, []

        # Exact match (case-insensitive)
        for candidate in candidate_services:
            if _normalize_name(candidate) == normalized:
                return candidate, sorted(candidate_services)

        # Match after stripping a leading "Azure " prefix
        if normalized.startswith("azure "):
            stripped = normalized.replace("azure ", "", 1)
            for candidate in candidate_services:
                if _normalize_name(candidate) == stripped:
                    return candidate, sorted(candidate_services)

        # Substring match
        matches = [
            candidate for candidate in candidate_services
            if normalized in _normalize_name(candidate)
            or _normalize_name(candidate) in normalized
        ]
        if matches:
            matches.sort(key=lambda name: len(name))
            return matches[0], sorted(candidate_services)

        return service, sorted(candidate_services)

    def _normalize_fabric_sku(self, sku: str | None) -> tuple[str | None, int | None]:
        """Normalize Fabric tier aliases (F2-F64) to CU-hour query."""
        if not sku:
            return None, None
        upper = sku.strip().upper()
        if upper in FABRIC_CAPACITY_TIERS:
            return "CU", FABRIC_CAPACITY_TIERS[upper]
        return sku, None

    def tool_azure_fetch_prices(
        self,
        services: list[str] | None = None,
        regions: list[str] | None = None,
        currency: str = "USD",
        force_refresh: bool = False,
    ) -> dict[str, Any]:
        """Fetch and cache Azure retail prices for specified services and regions.

        Downloads pricing data from the Azure Retail Prices API and stores it
        locally for fast subsequent queries. Data is cached on disk for 24 hours.

        Example:
            azure_fetch_prices(
                services=["Virtual Machines", "Azure Databricks", "Storage"],
                regions=["westeurope", "eastus"]
            )

            azure_fetch_prices(services=["API Management"], force_refresh=True)

        Args:
            services: List of Azure service names to fetch pricing for.
                     If None, fetches common services.
            regions: List of ARM region names (e.g., "westeurope", "eastus").
                    If None, fetches common regions.
            currency: Currency code (default: "USD")
            force_refresh: If True, ignore cache and fetch fresh data

        Returns:
            Dictionary with fetch status and summary statistics
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        # Default services commonly used in estimates
        if services is None:
            services = [
                "Virtual Machines",
                "Azure Databricks",
                "API Management",
                "Azure App Service",
                "Storage",
                "Azure Cosmos DB",
                "SQL Database",
                "ExpressRoute",
                "Azure Functions",
                "Azure Kubernetes Service",
            ]

        # Default regions
        if regions is None:
            regions = ["westeurope", "eastus", "westus2", "northeurope"]

        results = {
            "fetched": [],
            "cached": [],
            "errors": [],
            "total_items": 0,
            "timestamp": datetime.utcnow().isoformat() + "Z",
        }

        for service in services:
            for region in regions:
                cache_key = f"{service}_{region}_{currency}".lower().replace(" ", "-")
                cache_path = _get_cache_path(cache_key)

                # Check cache first
                if not force_refresh and _is_cache_valid(cache_path):
                    cached_data = _load_cache(cache_path)
                    if cached_data:
                        self._price_cache[cache_key] = cached_data.get("items", [])
                        results["cached"].append({
                            "service": service,
                            "region": region,
                            "items": len(cached_data.get("items", [])),
                        })
                        results["total_items"] += len(cached_data.get("items", []))
                        continue

                # Fetch from API
                try:
                    items = self._fetch_prices_from_api(service, region, currency)

                    # Cache to disk
                    cache_data = {
                        "service": service,
                        "region": region,
                        "currency": currency,
                        "fetched_at": datetime.utcnow().isoformat() + "Z",
                        "items": items,
                    }
                    _save_cache(cache_path, cache_data)

                    # Cache to memory
                    self._price_cache[cache_key] = items

                    results["fetched"].append({
                        "service": service,
                        "region": region,
                        "items": len(items),
                    })
                    results["total_items"] += len(items)

                except Exception as e:
                    results["errors"].append({
                        "service": service,
                        "region": region,
                        "error": str(e),
                    })

        return results

    def _fetch_prices_from_api(
        self, service: str, region: str, currency: str
    ) -> list[dict]:
        """Fetch all price items for a service/region from the API."""
        items = []

        filter_query = f"serviceName eq '{service}' and armRegionName eq '{region}'"
        url = f"{AZURE_PRICES_API}?currencyCode={currency}&$filter={quote(filter_query)}"

        max_retries = 3
        base_delay = 2  # seconds

        while url:
            for attempt in range(max_retries):
                try:
                    response = requests.get(url, timeout=30)

                    # Handle rate limiting
                    if response.status_code == 429:
                        if attempt < max_retries - 1:
                            delay = base_delay * (2 ** attempt)
                            time.sleep(delay)
                            continue
                        response.raise_for_status()

                    response.raise_for_status()
                    data = response.json()

                    items.extend(data.get("Items", []))
                    url = data.get("NextPageLink")
                    break  # Success, exit retry loop

                except requests.exceptions.RequestException:
                    if attempt < max_retries - 1:
                        delay = base_delay * (2 ** attempt)
                        time.sleep(delay)
                    else:
                        raise

        return items

    def tool_azure_query_prices(
        self,
        service: str | None = None,
        region: str = "westeurope",
        sku_contains: str | None = None,
        product_contains: str | None = None,
        price_type: str = "Consumption",
        currency: str = "USD",
        max_results: int = 50,
        page: int = 1,
    ) -> dict[str, Any]:
        """Query cached Azure pricing data with flexible filters.

        Searches the in-memory price cache for matching items. If data is not
        cached, it will be fetched automatically.

        Example:
            # Get VM pricing
            azure_query_prices(
                service="Virtual Machines",
                region="westeurope",
                sku_contains="D4"
            )

            # Get Databricks DBU pricing
            azure_query_prices(
                service="Azure Databricks",
                sku_contains="All-purpose"
            )

            # Get reserved instance pricing
            azure_query_prices(
                service="Virtual Machines",
                sku_contains="D4",
                price_type="Reservation"
            )

        Args:
            service: Azure service name (required for initial query)
            region: ARM region name (default: "westeurope")
            sku_contains: Filter by SKU name containing this string
            product_contains: Filter by product name containing this string
            price_type: Price type filter: "Consumption", "Reservation",
                       "DevTestConsumption", or None for all
            currency: Currency code (default: "USD")
            max_results: Maximum items to return (default: 50)

        Returns:
            Dictionary with matching price items and summary
        """
        if not service:
            return {"error": "service parameter is required"}

        resolved_service, candidates = self._resolve_service_name(
            service,
            region,
            currency,
            allow_fetch=True,
        )
        normalized_sku, fabric_cu = self._normalize_fabric_sku(sku_contains)

        cache_key = f"{resolved_service}_{region}_{currency}".lower().replace(" ", "-")

        # Load from memory cache or disk cache
        if cache_key not in self._price_cache:
            cache_path = _get_cache_path(cache_key)
            if cache_path.exists():
                cached_data = _load_cache(cache_path)
                if cached_data:
                    self._price_cache[cache_key] = cached_data.get("items", [])
            else:
                # Fetch if not cached
                fetch_result = self.tool_azure_fetch_prices(
                    services=[resolved_service], regions=[region], currency=currency
                )
                if "errors" in fetch_result and fetch_result["errors"]:
                    return {"error": fetch_result["errors"][0].get("error", "Fetch failed")}

        items = self._price_cache.get(cache_key, [])

        # Apply filters
        filtered = []
        for item in items:
            # Price type filter
            if price_type and item.get("type") != price_type:
                continue

            # SKU filter
            if normalized_sku:
                sku_name = item.get("skuName", "") or ""
                arm_sku = item.get("armSkuName", "") or ""
                if normalized_sku.lower() not in sku_name.lower() and \
                   normalized_sku.lower() not in arm_sku.lower():
                    continue

            # Product filter
            if product_contains:
                product_name = item.get("productName", "") or ""
                if product_contains.lower() not in product_name.lower():
                    continue

            filtered.append(item)

            if len(filtered) >= max_results * max(1, page):
                break

        # Format results
        start_index = max(0, (page - 1) * max_results)
        end_index = start_index + max_results
        paged_items = filtered[start_index:end_index]

        formatted_items = []
        for item in paged_items:
            formatted_items.append({
                "serviceName": item.get("serviceName"),
                "productName": item.get("productName"),
                "skuName": item.get("skuName"),
                "armSkuName": item.get("armSkuName"),
                "meterName": item.get("meterName"),
                "retailPrice": item.get("retailPrice"),
                "unitOfMeasure": item.get("unitOfMeasure"),
                "type": item.get("type"),
                "reservationTerm": item.get("reservationTerm"),
                "savingsPlan": item.get("savingsPlan"),
                "effectiveStartDate": item.get("effectiveStartDate"),
            })

        return {
            "service": resolved_service,
            "service_requested": service,
            "service_candidates": candidates[:50],
            "fabric_tier_cu": fabric_cu,
            "region": region,
            "price_type": price_type,
            "filters": {
                "sku_contains": sku_contains,
                "product_contains": product_contains,
            },
            "count": len(formatted_items),
            "truncated": len(filtered) > end_index,
            "next_page": page + 1 if len(filtered) > end_index else None,
            "page": page,
            "page_size": max_results,
            "items": formatted_items,
        }

    def tool_azure_calculate_cost(
        self,
        service: str,
        region: str = "westeurope",
        sku_match: str | None = None,
        product_match: str | None = None,
        quantity: int = 1,
        hours_per_month: int = 730,
        price_type: str = "Consumption",
        currency: str = "USD",
    ) -> dict[str, Any]:
        """Calculate monthly cost for an Azure resource.

        Looks up pricing and calculates the estimated monthly cost based on
        quantity and usage hours.

        Example:
            # Cost for 3 D4 v5 VMs running 24/7
            azure_calculate_cost(
                service="Virtual Machines",
                sku_match="D4 v5",
                quantity=3
            )

            # Cost for 1000 GB storage
            azure_calculate_cost(
                service="Storage",
                product_match="Blob Storage",
                sku_match="Hot LRS",
                quantity=1000,
                hours_per_month=1  # Storage is per GB, not per hour
            )

        Args:
            service: Azure service name
            region: ARM region name (default: "westeurope")
            sku_match: SKU name to match
            product_match: Product name to match
            quantity: Number of units (VMs, instances, GB, etc.)
            hours_per_month: Hours of usage per month (default: 730 = 24/7)
            price_type: Price type (default: "Consumption")
            currency: Currency code (default: "USD")

        Returns:
            Dictionary with pricing details and calculated costs
        """
        resolved_service, candidates = self._resolve_service_name(
            service,
            region,
            currency,
            allow_fetch=True,
        )

        normalized_sku, fabric_cu = self._normalize_fabric_sku(sku_match)
        effective_quantity = quantity
        if fabric_cu:
            effective_quantity = quantity * fabric_cu

        # Query matching prices
        result = self.tool_azure_query_prices(
            service=resolved_service,
            region=region,
            sku_contains=normalized_sku,
            product_contains=product_match,
            price_type=price_type,
            currency=currency,
            max_results=10,
        )

        if "error" in result:
            return result

        if not result.get("items"):
            return {
                "error": f"No pricing found for {resolved_service} with SKU '{sku_match}' in {region}",
                "suggestion": "Try broadening your search criteria",
            }

        # Find best match (first result)
        best_match = result["items"][0]
        unit_price = best_match.get("retailPrice", 0)
        unit_of_measure = best_match.get("unitOfMeasure", "")

        # Calculate cost based on unit of measure
        if "Hour" in unit_of_measure:
            monthly_cost = unit_price * effective_quantity * hours_per_month
            calculation = f"{unit_price} × {effective_quantity} × {hours_per_month} hours"
        elif "GB" in unit_of_measure:
            monthly_cost = unit_price * effective_quantity
            calculation = f"{unit_price} × {effective_quantity} GB"
        elif "10K" in unit_of_measure:
            monthly_cost = unit_price * (effective_quantity / 10000)
            calculation = f"{unit_price} × ({effective_quantity} / 10,000)"
        else:
            monthly_cost = unit_price * effective_quantity
            calculation = f"{unit_price} × {effective_quantity}"

        # Get savings plan info if available
        savings_info = None
        if best_match.get("savingsPlan"):
            savings_info = []
            for plan in best_match["savingsPlan"]:
                term = plan.get("term", "")
                plan_price = plan.get("retailPrice", 0)
                if "Hour" in unit_of_measure:
                    plan_monthly = plan_price * quantity * hours_per_month
                else:
                    plan_monthly = plan_price * quantity
                savings_pct = ((monthly_cost - plan_monthly) / monthly_cost * 100) if monthly_cost > 0 else 0
                savings_info.append({
                    "term": term,
                    "unit_price": plan_price,
                    "monthly_cost": round(plan_monthly, 2),
                    "savings_percent": round(savings_pct, 1),
                })

        return {
            "service": resolved_service,
            "service_requested": service,
            "service_candidates": candidates[:50],
            "region": region,
            "matched_sku": best_match.get("skuName"),
            "matched_product": best_match.get("productName"),
            "unit_price": unit_price,
            "unit_of_measure": unit_of_measure,
            "quantity": quantity,
            "effective_quantity": effective_quantity,
            "fabric_tier": sku_match if fabric_cu else None,
            "fabric_cu": fabric_cu,
            "hours_per_month": hours_per_month,
            "calculation": calculation,
            "monthly_cost": round(monthly_cost, 2),
            "annual_cost": round(monthly_cost * 12, 2),
            "currency": currency,
            "savings_plans": savings_info,
            "alternatives": [
                {
                    "sku": item.get("skuName"),
                    "product": item.get("productName"),
                    "unit_price": item.get("retailPrice"),
                }
                for item in result["items"][1:5]
            ] if len(result["items"]) > 1 else None,
        }

    def tool_azure_list_services(
        self,
        region: str = "westeurope",
        currency: str = "USD",
        from_cache_only: bool = True,
        max_pages: int = 2,
        max_services: int = 200,
    ) -> dict[str, Any]:
        """List available Azure service names for a region.

        Args:
            region: ARM region name (default: "westeurope")
            currency: Currency code (default: "USD")
            from_cache_only: If True, only uses cached data
            max_pages: Maximum API pages to scan when fetching (default: 2)
            max_services: Maximum services to return (default: 200)

        Returns:
            Dictionary with service names and source metadata
        """
        services = set()

        # Use memory cache
        for key, _items in self._price_cache.items():
            if key.endswith(f"_{region}_{currency}".lower().replace(" ", "-")):
                services.update(_extract_cache_service_names(_items))

        # Use disk cache
        if CACHE_DIR.exists():
            for cache_file in CACHE_DIR.glob("*.json"):
                cached = _load_cache(cache_file)
                if not cached:
                    continue
                if cached.get("region") != region or cached.get("currency") != currency:
                    continue
                services.update(_extract_cache_service_names(cached.get("items", [])))

        if services or from_cache_only:
            return {
                "region": region,
                "currency": currency,
                "services": sorted(services)[:max_services],
                "count": min(len(services), max_services),
                "source": "cache",
            }

        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        # Fetch limited pages for region to discover services
        filter_query = f"armRegionName eq '{region}'"
        url = f"{AZURE_PRICES_API}?currencyCode={currency}&$filter={quote(filter_query)}"
        pages = 0
        while url and pages < max_pages and len(services) < max_services:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            data = response.json()
            for item in data.get("Items", []):
                service_name = item.get("serviceName")
                if service_name:
                    services.add(service_name)
                    if len(services) >= max_services:
                        break
            url = data.get("NextPageLink")
            pages += 1

        return {
            "region": region,
            "currency": currency,
            "services": sorted(services)[:max_services],
            "count": min(len(services), max_services),
            "source": "api" if services else "none",
            "pages_scanned": pages,
        }

    def tool_azure_list_regions(
        self,
        service: str | None = None,
        currency: str = "USD",
        from_cache_only: bool = True,
        max_pages: int = 2,
        max_regions: int = 200,
    ) -> dict[str, Any]:
        """List available ARM regions for a service or cached data.

        Args:
            service: Optional service name to filter by
            currency: Currency code (default: "USD")
            from_cache_only: If True, only uses cached data
            max_pages: Maximum API pages to scan when fetching (default: 2)
            max_regions: Maximum regions to return (default: 200)

        Returns:
            Dictionary with region names and source metadata
        """
        regions = set()

        # Use memory cache
        for key, _items in self._price_cache.items():
            parts = key.rsplit("_", 2)
            if len(parts) >= 2:
                regions.add(parts[-2])

        # Use disk cache
        if CACHE_DIR.exists():
            for cache_file in CACHE_DIR.glob("*.json"):
                cached = _load_cache(cache_file)
                if not cached:
                    continue
                cached_region = cached.get("region")
                if cached_region:
                    regions.add(cached_region)

        if regions or from_cache_only:
            return {
                "regions": sorted(regions)[:max_regions],
                "count": min(len(regions), max_regions),
                "source": "cache",
            }

        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        filter_query = None
        if service:
            filter_query = f"serviceName eq '{service}'"
        url = f"{AZURE_PRICES_API}?currencyCode={currency}"
        if filter_query:
            url += f"&$filter={quote(filter_query)}"

        pages = 0
        while url and pages < max_pages and len(regions) < max_regions:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            data = response.json()
            for item in data.get("Items", []):
                region_name = item.get("armRegionName")
                if region_name:
                    regions.add(region_name)
                    if len(regions) >= max_regions:
                        break
            url = data.get("NextPageLink")
            pages += 1

        return {
            "regions": sorted(regions)[:max_regions],
            "count": min(len(regions), max_regions),
            "source": "api" if regions else "none",
            "pages_scanned": pages,
        }

    def tool_azure_list_cached_services(self) -> dict[str, Any]:
        """List all services currently cached in memory and on disk.

        Returns:
            Dictionary with cached service/region combinations and statistics
        """
        # Check memory cache
        memory_cache = []
        for key in self._price_cache:
            parts = key.rsplit("_", 2)
            if len(parts) >= 2:
                memory_cache.append({
                    "key": key,
                    "items": len(self._price_cache[key]),
                })

        # Check disk cache
        disk_cache = []
        if CACHE_DIR.exists():
            for cache_file in CACHE_DIR.glob("*.json"):
                try:
                    data = _load_cache(cache_file)
                    if data:
                        mtime = cache_file.stat().st_mtime
                        age_hours = (time.time() - mtime) / 3600
                        disk_cache.append({
                            "file": cache_file.name,
                            "service": data.get("service"),
                            "region": data.get("region"),
                            "items": len(data.get("items", [])),
                            "fetched_at": data.get("fetched_at"),
                            "age_hours": round(age_hours, 1),
                            "valid": age_hours < CACHE_TTL_HOURS,
                        })
                except Exception:
                    pass

        return {
            "memory_cache": {
                "entries": len(memory_cache),
                "items": memory_cache,
            },
            "disk_cache": {
                "directory": str(CACHE_DIR),
                "entries": len(disk_cache),
                "items": disk_cache,
            },
            "cache_ttl_hours": CACHE_TTL_HOURS,
        }

    def tool_azure_clear_cache(
        self,
        service: str | None = None,
        region: str | None = None,
        clear_disk: bool = True,
        clear_memory: bool = True,
    ) -> dict[str, Any]:
        """Clear Azure pricing cache.

        Args:
            service: Clear cache for specific service only (None = all)
            region: Clear cache for specific region only (None = all)
            clear_disk: Clear disk cache (default: True)
            clear_memory: Clear memory cache (default: True)

        Returns:
            Dictionary with cleared cache information
        """
        cleared = {"memory": [], "disk": []}

        # Build filter pattern
        if service and region:
            pattern = f"{service}_{region}".lower().replace(" ", "-")
        elif service:
            pattern = f"{service}_".lower().replace(" ", "-")
        elif region:
            pattern = f"_{region}_".lower()
        else:
            pattern = None

        # Clear memory cache
        if clear_memory:
            keys_to_remove = []
            for key in self._price_cache:
                if pattern is None or pattern in key:
                    keys_to_remove.append(key)
            for key in keys_to_remove:
                del self._price_cache[key]
                cleared["memory"].append(key)

        # Clear disk cache
        if clear_disk and CACHE_DIR.exists():
            for cache_file in CACHE_DIR.glob("*.json"):
                if pattern is None or pattern in cache_file.stem:
                    cache_file.unlink()
                    cleared["disk"].append(cache_file.name)

        return {
            "cleared_memory": len(cleared["memory"]),
            "cleared_disk": len(cleared["disk"]),
            "details": cleared,
        }
