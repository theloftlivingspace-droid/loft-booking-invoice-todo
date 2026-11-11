/**
 * FixApartmenteryCustomerHuangfuZifu20261111.gs
 * -----------------------------------------------------------------------
 * Follow-up to RenameGuestBaiBaiToHuangfuZifu20261111.gs. Apartmentery
 * booking 339270 (room 210, 10→20 Nov 2026) is attached to an EXISTING
 * customer record "Bai Bai / Booking" (customerType=existing + customerId).
 * For existing customers the server ignores the customerName field, so the
 * earlier edit redirected "OK" but the displayed name never changed.
 *
 * This re-submits the booking edit form pointing at a customer named
 * "Huangfu Zifu / Booking": reuses an existing customer with that exact
 * name if there is one, otherwise sends customerType=new so Apartmentery
 * creates it. Every other field is re-sent untouched. Verifies afterwards
 * by re-reading the form. The old "Bai Bai / Booking" customer record is
 * NOT deleted (delete it in Apartmentery by hand if you want).
 *
 * HOW TO RUN: Apps Script editor ▶ fixApartmenteryCustomerHuangfuZifu20261111 ▶ Run
 */

function fixApartmenteryCustomerHuangfuZifu20261111() {
  const BOOKING_ID = '339270';
  const ROOM = '210';
  const NEW_CUSTOMER = 'Huangfu Zifu / Booking';

  const unit = getApartmenteryUnitForRoom(ROOM);
  if (!unit) { Logger.log('ABORT — no unit for room ' + ROOM); return; }
  const path = `/user/branch/${unit.branchId}/unit/${unit.unitId}/booking/${BOOKING_ID}/edit`;

  const html = _apartmenteryFetch_(path, { method: 'get' }).getContentText();
  const before = _selectedCustomerLabel20261111_(html);
  Logger.log(`Before: customerType=${_extractCheckedRadioValue_(html, 'customerType')}, customer="${before.label}" (id ${before.id})`);

  const state = _getApartmenteryBookingEditFormState_(unit.branchId, unit.unitId, BOOKING_ID);
  const found = _findExistingApartmenteryCustomerId_(html, NEW_CUSTOMER);
  if (found.status === 'ambiguous') { Logger.log('ABORT — 2+ customers already named like that: ' + found.ids.join(',')); return; }

  const payload = {
    startDate: state.startDate,
    endDate: state.endDate,
    note: state.note,
    customerName: NEW_CUSTOMER,
    customerMobileNo: state.customerMobileNo,
    customerIdNo: state.customerIdNo,
    customerEmail: state.customerEmail,
    customerNote: state.customerNote,
    remindEvery: state.remindEvery,
    reminderFrequency: state.reminderFrequency,
    remindOnDayInMonth: state.remindOnDayInMonth,
    remindOnDayInWeek: state.remindOnDayInWeek,
    remindOnDayInMonthInYear: state.remindOnDayInMonthInYear,
    remindOnMonthInYear: state.remindOnMonthInYear,
    remindInvoiceDayBefore: state.remindInvoiceDayBefore
  };
  if (state.reminderChecked) payload.reminder = 'true';
  if (found.status === 'found') {
    payload.customerType = 'existing';
    payload.customerId = found.id;
    Logger.log('Re-linking to existing customer id ' + found.id);
  } else {
    payload.customerType = 'new';
    Logger.log('Creating new customer "' + NEW_CUSTOMER + '"');
  }

  const resp = _apartmenteryFetch_(path, { method: 'post', payload: payload });
  const code = resp.getResponseCode();
  if (!(code >= 300 && code < 400)) {
    Logger.log(`FAILED — HTTP ${code}: ${_extractPlayErrorMessage_(resp.getContentText())}`);
    return;
  }

  const after = _selectedCustomerLabel20261111_(_apartmenteryFetch_(path, { method: 'get' }).getContentText());
  Logger.log(`After: customer="${after.label}" (id ${after.id})`);
  Logger.log(/huangfu/i.test(after.label) ? 'DONE — name changed ✔' : 'WARNING — redirect OK but the customer still reads "' + after.label + '"; change it by hand.');
}

function _selectedCustomerLabel20261111_(html) {
  const sel = html.match(/<select[^>]*id="customerId"[\s\S]*?<\/select>/);
  if (!sel) return { id: '', label: '(no customer select found)' };
  const m = sel[0].match(/<option\s+value="([^"]*)"[^>]*selected[^>]*>([\s\S]*?)<\/option>/);
  return m ? { id: m[1], label: m[2].replace(/\s+/g, ' ').trim() } : { id: '', label: '(none selected)' };
}
