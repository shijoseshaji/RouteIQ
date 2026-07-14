# 🗺️ RouteIQ
### Route Intelligence & Optimization Platform

RouteIQ is a first mile pickup route optimizer for logistics operations.

## What it does
- Assigns pickup executives to sellers based on geography
- Optimizes routes using Nearest Neighbour + 2-opt algorithm
- Distributes sellers across cut-off slots (12PM, 2PM, 4PM, 6PM)
- Auto-splits executive routes if KM or working hours exceed limits
- Generates full Excel report with transit times and distances

## Input Required
Excel file with two sheets:

**Sellers sheet**
- seller_id
- volume
- latitude
- longitude
- hub_id

**Hubs sheet**
- hub_id
- hub_lat
- hub_lon

## Configuration
All settings adjustable from the sidebar:
- Shift start time
- Max working hours
- Max route KM
- Min productivity
- Travel speed
- Pickup time per shipment

## Output
- Seller Assignments with cut-off slots
- Executive Summary with KM and hours validation
- Slot Distribution
- Transit Detail (inter-seller legs)
- Route Summary
- Split Log (if any routes exceeded limits)
- Outliers flagged for manual review
