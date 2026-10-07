"""Command line: python -m comps AMZN [--peers 5] [--with MSFT,GOOGL] [--out file.xlsx]"""
import argparse
import sys

from .data import DataError
from .excel import build_workbook
from .pipeline import analyze


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="python -m comps", description="Automated comparable-company analysis")
    ap.add_argument("ticker")
    ap.add_argument("--peers", type=int, default=5, help="number of peers to auto-select (default 5)")
    ap.add_argument("--with", dest="manual", default="", help="comma-separated peer tickers to use instead")
    ap.add_argument("--out", default=None, help="Excel output path (default <TICKER>_comps.xlsx)")
    ap.add_argument("--sample", action="store_true", help="use bundled offline sample data (AMZN only)")
    args = ap.parse_args(argv)

    provider = None
    if args.sample:
        from .sample import SampleProvider
        provider = SampleProvider()
    manual = [t.strip().upper() for t in args.manual.split(",") if t.strip()]
    try:
        a = analyze(args.ticker, provider=provider, n_peers=args.peers, manual_peers=manual or None)
    except DataError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    r = a.result
    print(f"\n{r.target.name} ({r.target.ticker})  {r.target.currency} {r.target.price:,.2f}")
    print(f"Profile: {r.selection.profile} – {r.selection.explanation}\n")
    print("Peers: " + ", ".join(f"{s.ticker} ({s.score:.0f})" for s in a.candidates[:len(r.peers)]))
    print(f"\n{'Multiple':<20}{'Target':>9}{'Median':>9}{'Prem.':>9}{'Implied':>11}")
    for m in r.multiples:
        fmt = lambda v, s="x": f"{v:.1f}{s}" if v is not None else "NM"
        prem = f"{m.premium:+.0%}" if m.premium is not None else "NM"
        imp = f"{m.implied_mid:,.2f}" if m.implied_mid is not None else "NM"
        print(f"{m.label + (' *' if m.is_primary else ''):<20}{fmt(m.target):>9}{fmt(m.median):>9}{prem:>9}{imp:>11}")
    print(f"\n{r.verdict}\n")

    out = args.out or f"{r.target.ticker}_comps.xlsx"
    build_workbook(r, a.candidates, query=args.ticker).save(out)
    print(f"Excel workbook saved to {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
