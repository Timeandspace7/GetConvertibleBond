"""Deterministic stock and bond allocation scenario calculator."""
import json
import sys
from decimal import Decimal, ROUND_FLOOR, ROUND_CEILING

def d(value):
    x = Decimal(str(value))
    if not x.is_finite():
        raise ValueError("Inputs must be finite")
    return x

def calculate(data):
    n, pb, q, bb = (d(data[k]) for k in ("shares", "buy_price", "bond_quantity", "bond_buy_price"))
    f = d(0 if data.get("total_fees") is None else data["total_fees"])
    bf, tick, offset = d(data.get("buy_fees", 0)), d(data.get("tick", .01)), d(data.get("offset", .1))
    if n <= 0 or n != n.to_integral_value() or q < 0 or q != q.to_integral_value():
        raise ValueError("Invalid share or bond quantity")
    if min(pb, bb) <= 0 or min(f, bf, offset) < 0 or tick <= 0:
        raise ValueError("Invalid prices, fees, tick or offset")
    if data.get("total_fees") is not None and bf > f:
        raise ValueError("Total fees cannot be less than buy fees")
    capital = n * pb + q * bb + bf
    rows = []
    for bs in map(d, data.get("bond_prices", [95, 100, 110, 120, 145])):
        if bs < 0:
            raise ValueError("Negative bond price")
        bond = q * (bs - bb)
        be = pb - (bond - f) / n
        candidates = [("理论平衡", be), ("上方情景", be + offset), ("下方情景", be - offset),
                      ("相邻报价下", (be / tick).to_integral_value(rounding=ROUND_FLOOR) * tick),
                      ("相邻报价上", (be / tick).to_integral_value(rounding=ROUND_CEILING) * tick)]
        seen = set()
        for kind, ps in candidates:
            if ps < 0 or ps in seen:
                continue
            seen.add(ps)
            stock = n * (ps - pb)
            total = stock + bond - f
            rows.append(dict(bond_price=float(bs), stock_price=float(ps), kind=kind, stock_pnl=float(stock),
                             bond_pnl=float(bond), fees=float(f), total_pnl=float(total),
                             return_rate=float(total/capital), breakeven=float(be)))
    return dict(version="1.0.1", nominal_capital=float(capital),
                fee_mode="毛盈亏：未计费用" if data.get("total_fees") is None else "净盈亏：固定费用假设", rows=rows)

def self_test():
    base = dict(shares=600, buy_price=4.24, bond_quantity=10, bond_buy_price=100, total_fees=None, bond_prices=[120])
    a = calculate(base)
    assert abs(a["rows"][0]["breakeven"] - 3.906666666666667) < 1e-12
    assert abs(a["rows"][0]["total_pnl"]) < 1e-10
    assert abs(a["rows"][1]["total_pnl"] - 60) < 1e-10
    assert calculate(dict(base, bond_quantity=0))["rows"][0]["breakeven"] == 4.24
    assert calculate(dict(base, bond_quantity=20))["nominal_capital"] == 4544
    assert abs(calculate(dict(base, total_fees=12))["rows"][0]["breakeven"] - 3.926666666666667) < 1e-12
    for patch in [dict(shares=0), dict(bond_quantity=-1), dict(tick=0), dict(buy_price="NaN")]:
        try:
            calculate(dict(base, **patch))
        except ValueError:
            pass
        else:
            raise AssertionError(patch)
    print("Scenario and boundary checks passed")

if __name__ == "__main__":
    if sys.argv[1:] == ["--self-test"]:
        self_test()
    else:
        with open(sys.argv[1], encoding="utf-8-sig") as source:
            print(json.dumps(calculate(json.load(source)), ensure_ascii=False, indent=2))
