/**
 * FixDuplicateInvoiceKirillBogdanov20260820.gs
 * -----------------------------------------------------------------------
 * One-off fix for the single case scanForDuplicateAmountInvoicesLive()
 * flagged 2026-08-20: booking 328625 (room 204, Kirill Bogdanov) had two
 * Apartmentery invoices for the exact same amount (1,583.75) — invoice
 * 2808679 (tracked under invoiceKey "ABB-HMPPESQ5SC", the Airbnb-parsed
 * row) and invoice 2809120 (tracked under invoiceKey
 * "SCB-2026-08-01-1583.75", the SCB-bank-matched row). Both rows represent
 * the SAME real payout — payout-income-log's Airbnb-email pipeline and its
 * SCB-statement-matching pipeline each independently produced a row for it
 * with no dedup between the two, so autoCreateApartmenteryInvoicesAndReceipts
 * created an invoice for each. Root cause lives in payout-income-log, not
 * here — this file only repairs the resulting bookkeeping in THIS repo.
 *
 * Nathan confirmed both invoices had receipts already and chose to keep
 * 2809120, manually deleting invoice 2808679 (and its receipt) directly on
 * Apartmentery.
 *
 * This script does NOT touch Apartmentery — it only repoints this repo's
 * local invoice_apt_ids_v1 record for invoiceKey "ABB-HMPPESQ5SC" away
 * from the now-deleted 2808679 to the surviving 2809120, so the-loft-admin's
 * 🧾 link for that row resolves to a real invoice instead of a 404. Leaves
 * invoice_done_v1 untouched (already true for this key — correctly so,
 * since a real invoice for this payout does exist, just under a different
 * invoiceId now).
 *
 * HOW TO RUN: Apps Script editor ▶ fixDuplicateInvoiceKirillBogdanov20260820
 * ▶ Run ▶ read log. Safe to re-run (idempotent — just re-writes the same
 * mapping).
 */
function fixDuplicateInvoiceKirillBogdanov20260820() {
  const invoiceKey = 'ABB-HMPPESQ5SC';
  const aptBookingId = '328625';
  const oldInvoiceId = '2808679'; // deleted manually on Apartmentery by Nathan
  const survivingInvoiceId = '2809120';

  const map = getProp_(PROP_KEY_INVOICE_APT_IDS);
  const before = map[invoiceKey] || '(none)';

  setInvoiceApartmenteryIds(invoiceKey, aptBookingId, survivingInvoiceId);

  Logger.log(`fixDuplicateInvoiceKirillBogdanov20260820: ${invoiceKey} "${before}" -> ` +
    `"${aptBookingId}:${survivingInvoiceId}" (deleted duplicate was ${oldInvoiceId})`);

  return { invoiceKey: invoiceKey, before: before, after: aptBookingId + ':' + survivingInvoiceId };
}
