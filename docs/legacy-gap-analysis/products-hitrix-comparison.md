# HITRIX `tblMastItem` vs Dhanman `products` Comparison

## Purpose

This document compares the current Dhanman `public.products` table with the legacy HITRIX item master usage visible in `hitrix-srk/src`.

It is focused on product master parity, missing parent dependencies, and recommended improvements for safer day-to-day operations.

## Current Dhanman Product Scope

Dhanman `public.products` currently includes:

- product identity: `id`, `product_name`, `product_code`, `sku`, `barcode`
- product relationships: `vendor_id`, `unit_id`, `base_unit_id`, `purchase_unit_id`, `issue_unit_id`, `default_purchase_unit_id`, `default_sales_unit_id`, `category_id`, `tax_category_id`, `product_group_id`
- product description and tax metadata: `description`, `hsn_code`, `sac`, `stock_type`, `tracking_type`
- default commercial values: `purchase_price`, `selling_price`
- stock controls: `opening_stock_base_qty`, `minimum_stock`, `maximum_stock`, `reorder_level`, `reorder_quantity`
- lifecycle fields: `is_active`, `is_stock_item`, `is_deleted`, audit timestamps, audit users

## HITRIX Evidence Visible In Source

The legacy codebase shows strong usage of `tblMastItem` for:

- item code and item name
- HSN display and GST summary support
- mill-linked item lookup
- standard pack and standard rate-per usage
- brokerage logic through item fields such as `Itbrokrt` and `ItbrokOn`
- stock control linkage through item control numbers and stock routines

Visible code evidence includes:

- `hitrix-srk/src/Access/DataCollection.bas`
- `hitrix-srk/src/Access/Stock.bas`
- `hitrix-srk/src/Access/frmSales GST.frm`
- `hitrix-srk/src/Access/frmPurchaseTrade.frm`
- `hitrix-srk/src/Access/frmBookingInward.frm`

## Comparison Matrix

| Area | HITRIX `tblMastItem` / behavior | Dhanman `products` | Status | Notes |
| --- | --- | --- | --- | --- |
| Core identity | Item code and item name are clearly present | `id`, `product_name`, `product_code`, `sku`, `barcode` | `Present` | Dhanman has stronger modern identity options. |
| HSN / tax metadata | HSN is visible in sales and GST reporting usage | `hsn_code`, `sac`, `tax_category_id` | `Present` | Good parity anchor. |
| Unit relationships | Item setup appears tied to standard pack/rate and transaction behavior | `unit_id`, `base_unit_id`, `purchase_unit_id`, `issue_unit_id`, defaults | `Present` | Dhanman is structurally richer here. |
| Category / grouping | Legacy grouping exists but exact item-group mapping needs more evidence | `category_id`, `product_group_id` | `Partial` | Dhanman grouping is present; legacy field mapping still needs verification. |
| Default prices | Legacy standard rate behavior is visible | `purchase_price`, `selling_price` | `Present` | Dhanman supports default price storage. |
| Brokerage behavior | Legacy item master appears to carry brokerage attributes | No explicit brokerage fields in `products` | `Partial` | If brokerage is still a business need, it likely needs separate modeling. |
| Stock controls | Legacy stock routines heavily reference item master and item control logic | opening, min/max stock, reorder controls, tracking type | `Present` | Dhanman supports stock-oriented controls, though not the same legacy pattern. |
| Parent master dependencies | Legacy item behavior depends on linked mill/account and other lookups | vendor, unit, category, tax category, group references exist | `Present` | Dhanman has clear parent references. |
| End-user price guardrails | No exact field match proven yet from visible legacy source | No `MinSellingPrice` or `MaxPurchasePrice` fields | `Missing` | Recommended as a Dhanman improvement to reduce pricing mistakes. |

## Missing Parent / Master Considerations

The Dhanman product model already depends on several parent masters and should keep them mandatory or well-governed where business rules require them:

- `units`
- `categories`
- `product_group`
- `product_tax_categories`
- `users`
- vendor/account master relation via `vendor_id`

For migration and validation, these parent dependencies should be checked before product import or user entry.

## Recommended Improvement

### Add Price Guardrails

To protect end users from accidental pricing mistakes, Dhanman should add:

- `MinSellingPrice decimal(18,2)`
- `MaxPurchasePrice decimal(18,2)`

### Why This Helps

- prevents users from selling below an approved floor price
- prevents users from purchasing above an approved ceiling price
- reduces manual mistakes during invoice or bill entry
- provides a simple control without changing the default `selling_price` and `purchase_price` model

### Recommended Validation Rules

- In sales flows, block or warn when line price is below `MinSellingPrice`
- In purchase flows, block or warn when line price is above `MaxPurchasePrice`
- Allow privileged override only through approval or elevated permission if business requires it
- Default both fields to `0` so legacy or unrestricted products continue to work

## Recommendation Summary

| Recommendation | Priority | Reason |
| --- | --- | --- |
| Add `MinSellingPrice` to product master | `High` | Protects against underpriced sales |
| Add `MaxPurchasePrice` to product master | `High` | Protects against overpriced purchases |
| Add validation in sales and purchase entry | `High` | Field addition alone is not enough |
| Confirm whether HITRIX had an equivalent rule | `Medium` | Useful for parity documentation, but not required to justify the safeguard |

## Conclusion

Dhanman already has a strong modern product master, but it does not currently protect users from entering unsafe sales or purchase prices at the item level.

Even though exact legacy evidence for these two fields is not yet proven from the visible HITRIX source, adding `MinSellingPrice` and `MaxPurchasePrice` is a sound improvement and should be treated as a high-priority operational safeguard.
