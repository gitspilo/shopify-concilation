"""
operation_cost.py — Per-parcel operational cost calculator (multi-brand).

Usage:
    python operation_cost.py

Prompts for fixed costs and brand order counts, then splits costs
proportionally and shows per-parcel cost for each brand.
"""

def ask_float(prompt, default=None):
    while True:
        suffix = f" [{default}]" if default is not None else ""
        val = input(f"{prompt}{suffix}: ").strip()
        if val == "" and default is not None:
            return float(default)
        try:
            return float(val)
        except ValueError:
            print("  Please enter a valid number.")

def ask_int(prompt, default=None):
    while True:
        suffix = f" [{default}]" if default is not None else ""
        val = input(f"{prompt}{suffix}: ").strip()
        if val == "" and default is not None:
            return int(default)
        try:
            return int(val)
        except ValueError:
            print("  Please enter a valid whole number.")

def run():
    print()
    print("=" * 60)
    print("   OPERATIONAL COST CALCULATOR")
    print("=" * 60)

    # ── Fixed costs ──────────────────────────────────────────────
    print("\n── Fixed Costs (monthly) ──")
    rent       = ask_float("  Rent",       51000)
    lightbill  = ask_float("  Light Bill", 15000)
    salary     = ask_float("  Salary",     50000)

    fixed_costs = {
        "Rent":       rent,
        "Light Bill": lightbill,
        "Salary":     salary,
    }

    # ── Brands ───────────────────────────────────────────────────
    print("\n── Brands & Dispatched Orders ──")
    brands = {}
    print("  Enter brand names and order counts.")
    print("  Press Enter with blank name when done.\n")

    # Pre-fill last known values as defaults
    defaults = [("KUDBI", 2496), ("SOFTSTITCH", 198)]
    idx = 0
    while True:
        default_name = defaults[idx][0] if idx < len(defaults) else ""
        name_suffix  = f" [{default_name}]" if default_name else ""
        name = input(f"  Brand name{name_suffix}: ").strip()
        if name == "":
            name = default_name
        if name == "":
            if len(brands) >= 1:
                break
            print("  At least one brand is required.")
            continue
        default_orders = defaults[idx][1] if idx < len(defaults) else None
        orders = ask_int(f"  Orders dispatched for {name}", default_orders)
        brands[name] = orders
        idx += 1

    # ── Calculation ──────────────────────────────────────────────
    total_orders = sum(brands.values())
    total_fixed  = sum(fixed_costs.values())
    brand_names  = list(brands.keys())
    col_w = 16

    print()
    print("=" * (14 + col_w * (1 + len(brand_names))))
    print(f"  COST SPLIT  |  Total orders: {total_orders}")
    print("=" * (14 + col_w * (1 + len(brand_names))))

    # Header row
    print(f"{'Cost Head':<14}", end="")
    print(f"{'Total':>{col_w}}", end="")
    for b in brand_names:
        print(f"{b:>{col_w}}", end="")
    print()
    print("-" * (14 + col_w * (1 + len(brand_names))))

    brand_totals = {b: 0.0 for b in brand_names}

    for cost_name, amount in fixed_costs.items():
        print(f"{cost_name:<14}", end="")
        print(f"{amount:>{col_w},.2f}", end="")
        for b in brand_names:
            share = amount * brands[b] / total_orders
            brand_totals[b] += share
            print(f"{share:>{col_w},.2f}", end="")
        print()

    # Total row
    print("-" * (14 + col_w * (1 + len(brand_names))))
    print(f"{'TOTAL':<14}", end="")
    print(f"{total_fixed:>{col_w},.2f}", end="")
    for b in brand_names:
        print(f"{brand_totals[b]:>{col_w},.2f}", end="")
    print()

    # Per-parcel row
    print("-" * (14 + col_w * (1 + len(brand_names))))
    print(f"{'Per Parcel':<14}", end="")
    print(f"{'':>{col_w}}", end="")
    for b in brand_names:
        per_parcel = brand_totals[b] / brands[b]
        print(f"{per_parcel:>{col_w},.2f}", end="")
    print()

    print("=" * (14 + col_w * (1 + len(brand_names))))
    print()
    print("  Share basis:")
    for b in brand_names:
        print(f"    {b}: {brands[b]} orders ({brands[b]/total_orders*100:.1f}%)")
    print()

if __name__ == "__main__":
    run()
