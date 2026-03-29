from ingestion import PathFinderParser
import os

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()

print(f"Grant Active: {data.grant_params.active}")
print(f"Grant Rate: {data.grant_params.rate}")
print(f"Grant Cap: {data.grant_params.cap}")

print(f"CCfD Active: {data.ccfd_params.active}")
print(f"CCfD Duration: {data.ccfd_params.duration}")
print(f"CCfD Type: {data.ccfd_params.contract_type}")
print(f"CCfD EUA Price Pct (Tau): {data.ccfd_params.eua_price_pct}")

years = sorted(data.time_series.carbon_prices.keys())
print("\nCarbon Prices:")
for y in years[:10]:
    print(f"  {y}: {data.time_series.carbon_prices[y]}")
