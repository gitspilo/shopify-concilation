"""
operation_cost.py — Per-parcel operational cost calculator (multi-brand).

Usage:
    python operation_cost.py

Prompts for fixed costs and, per brand, the orders dispatched and the
shipments delivered. Costs are split by dispatched share (an RTO parcel
costs the same to pack and ship as a delivered one), then shown both per
dispatched parcel and per delivered parcel.
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


def ask_brand(name, default_orders=None, default_delivered=None):
    """Dispatched and delivered counts for one brand.

    Delivered can't exceed dispatched, so re-ask instead of letting a typo
    through into a per-delivered figure that reads lower than it should.
    """
    while True:
        orders = ask_int(f"    Orders dispatched for {name}", default_orders)
        if orders < 0:
            print("    Can't be negative.")
            continue
        delivered = ask_int(f"    Shipments delivered for {name}", default_delivered)
        if delivered < 0:
            print("    Can't be negative.")
            continue
        if delivered > orders:
            print(f"    Delivered ({delivered}) can't exceed dispatched ({orders}).")
            continue
        return {"orders": orders, "delivered": delivered}


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
    print("\n── Brands: Dispatched & Delivered ──")
    print("  Each known brand is offered in turn; answer n to leave one out.")
    print("  After those, you're asked whether to add any further brands.")
    print("  Both counts come from the reconcile.py summary: fulfilled orders,")
    print("  and Bluedart delivered + Other delivered.\n")
    brands = {}

    # Known brands, with last-run counts as defaults: (dispatched, delivered)
    known = [
        ("KUDBI",      2857, 1819),
        ("SOFTSTITCH",  198, None),
        ("KANTHAKALA",   32,   16),
    ]
    for name, default_orders, default_delivered in known:
        if ask_yes_no(f"  Include {name}?"):
            brands[name] = ask_brand(name, default_orders, default_delivered)

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
        brands[name] = ask_brand(name)

    # ── Calculation ──────────────────────────────────────────────
    # Costs split by DISPATCHED share: an RTO parcel costs the same to pack,
    # store and hand to the courier as one that lands.
    total_orders    = sum(b["orders"] for b in brands.values())
    total_delivered = sum(b["delivered"] for b in brands.values())
    if total_orders <= 0:
        print("\n  Total dispatched orders is 0 — nothing to split."
              " Re-run with real counts.")
        return
    total_fixed  = sum(fixed_costs.values())
    brand_names  = list(brands.keys())
    col_w = 16
    width = 14 + col_w * (1 + len(brand_names))

    def row(label, total, values, fmt=",.2f"):
        print(f"{label:<14}", end="")
        print(f"{total:>{col_w}{fmt}}" if total is not None else f"{'':>{col_w}}", end="")
        for v in values:
            print(f"{v:>{col_w}{fmt}}" if v is not None else f"{'—':>{col_w}}", end="")
        print()

    print()
    print("=" * width)
    print(f"  COST SPLIT  |  Dispatched: {total_orders}  |  Delivered: {total_delivered}")
    print("=" * width)

    # Header row
    print(f"{'Cost Head':<14}", end="")
    print(f"{'Total':>{col_w}}", end="")
    for b in brand_names:
        print(f"{b:>{col_w}}", end="")
    print()
    print("-" * width)

    # Volume rows
    row("Dispatched", total_orders, [brands[b]["orders"] for b in brand_names], fmt=",d")
    row("Delivered", total_delivered, [brands[b]["delivered"] for b in brand_names], fmt=",d")
    row("Delivery%",
        total_delivered / total_orders * 100,
        [brands[b]["delivered"] / brands[b]["orders"] * 100 if brands[b]["orders"] else None
         for b in brand_names])
    print("-" * width)

    brand_totals = {b: 0.0 for b in brand_names}

    for cost_name, amount in fixed_costs.items():
        shares = []
        for b in brand_names:
            share = amount * brands[b]["orders"] / total_orders
            brand_totals[b] += share
            shares.append(share)
        row(cost_name, amount, shares)

    print("-" * width)
    row("TOTAL", total_fixed, [brand_totals[b] for b in brand_names])

    # Per-parcel rows. Per-dispatched is the same for every brand by
    # construction (cost split by order share, divided by that share again);
    # per-delivered is the one that varies — an RTO-heavy brand carries the
    # cost of parcels that earned nothing.
    print("-" * width)
    row("Per Dispatch", None,
        [brand_totals[b] / brands[b]["orders"] if brands[b]["orders"] else None
         for b in brand_names])
    row("Per Delivered", None,
        [brand_totals[b] / brands[b]["delivered"] if brands[b]["delivered"] else None
         for b in brand_names])
    print("=" * width)

    print()
    print("  Share basis (dispatched orders):")
    for b in brand_names:
        d = brands[b]
        pct = d["orders"] / total_orders * 100
        print(f"    {b}: {d['orders']} dispatched ({pct:.1f}%),"
              f" {d['delivered']} delivered")
    zero = [b for b in brand_names if not brands[b]["delivered"]]
    if zero:
        print(f"\n  No delivered shipments for {', '.join(zero)}"
              f" — per-delivered shown as '—'.")
    print()

if __name__ == "__main__":
    run()
