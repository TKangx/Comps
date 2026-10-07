"""Comps Analyzer: automated comparable-company valuation from a single ticker."""
from .analysis import CompsResult, run_comps
from .data import CompanyData, DataError, YahooProvider
from .peers import PeerScore, find_peers, score_manual_peers
from .pipeline import analyze

__all__ = ["analyze", "run_comps", "find_peers", "score_manual_peers", "CompanyData", "CompsResult",
           "PeerScore", "YahooProvider", "DataError"]
