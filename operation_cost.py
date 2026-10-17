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

def ask_yes_no(prompt, default=True):
    suffix = " [Y/n]" if default else " [y/N]"
    while True:
        val = input(f"{prompt}{suffix}: ").strip().lower()
        if val == "":
            return default
        if val in ("y", "yes"):
            return True
        if val in ("n", "no"):
            return False
        print("  Please answer y or n.")


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
    print("  Each known brand is offered in turn; answer n to leave one out.")
    print("  After those, you're asked whether to add any further brands.\n")
    brands = {}

    # Known brands, with last-run order counts as defaults
    known = [("KUDBI", 2857), ("SOFTSTITCH", 198), ("KANTHAKALA", 32)]
    for name, default_orders in known:
        if ask_yes_no(f"  Include {name}?"):
            brands[name] = ask_int(f"    Orders dispatched for {name}", default_orders)

    # Any brand not in the list above
    while True:
        if not ask_yes_no("\n  Add another brand?", default=False):
            if brands:
                break
            print("  At least one brand is required.")
            continue
        name = input("    Brand name: ").strip()
        if not name:
            print("    Brand name can't be blank.")
            continue
        if name in brands:
            print(f"    {name} is already entered.")
            continue
        brands[name] = ask_int(f"    Orders dispatched for {name}")

    # ── Calculation ──────────────────────────────────────────────
    total_orders = sum(brands.values())
    if total_orders <= 0:
        print("\n  Total orders is 0 — nothing to split. Re-run with real counts.")
        return
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
